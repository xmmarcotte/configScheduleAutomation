import sys
import textwrap
import requests
import time
from smartsheet import smartsheet
import smartsheet, smartsheet.exceptions
import pymssql
from pymssql import Error
import usaddress
from geopy.geocoders import Nominatim
import json
import re
import os
from smartsheet import fresh_operation
import datetime
import random
import logging
import pytz
import calendar
from dotenv import load_dotenv
from typing import Dict, Any, List

today = datetime.datetime.today().date()
filename = f"C:\\Users\\mmarcotte\\Documents\\Smartsheet Automation Logfiles\\{today}.log"
root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(filename)
fh.setLevel(logging.INFO)

# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.WARNING)
# create formatter and add it to the handlers
ch_formatter = logging.Formatter('%(message)s')
fh_formatter = logging.Formatter('%(asctime)s.%(msecs)03d |    %(message)s', '%Y-%m-%d %H:%M:%S')
ch.setFormatter(ch_formatter)
fh.setFormatter(fh_formatter)
# add the handlers to logger
root_logger.addHandler(ch)
root_logger.addHandler(fh)
load_dotenv()


def exponential_backoff(attempt, max_attempts=5, base_delay=60, max_delay=300):
    if attempt >= max_attempts:
        return False
    # calculate delay, adding a small random factor to avoid congestion
    delay = min(max_delay, base_delay * 2 ** attempt) + random.uniform(0, 10)
    print(f"waiting for {delay:.2f} seconds before retrying...")
    time.sleep(delay)
    return True


def smartsheet_api_call_with_retry(call, *args, **kwargs):
    attempt = 0
    max_attempts = 5
    while attempt < max_attempts:
        try:
            # make the smartsheet sdk call
            return call(*args, **kwargs)
        except smartsheet.exceptions.ApiError as e:
            # rate limiting error, use backoff
            if e.error.result.error_code == 4003:
                print("encountered rate limit error, applying exponential backoff.")
                if not exponential_backoff(attempt, max_attempts):
                    print("max retry attempts reached for rate limit error. giving up.")
                    return None
            # handle 5XX server errors with shorter backoff
            elif 500 <= e.error.result.status_code < 600:
                print(f"encountered server error with status code: {e.error.result.status_code}, applying short delay.")
                if not exponential_backoff(attempt, max_attempts, base_delay=10, max_delay=60):
                    print("max retry attempts reached for server error. giving up.")
                    return None
            # handle other api errors
            else:
                print(f"smartsheet api error: {e}")
                return None
        except Exception as e:
            # unexpected error, not retrying
            print(f"unexpected error: {e}")
            return None
        attempt += 1


def normalize_ticket_number(ticket_num):
    ticket_str = str(ticket_num).strip()

    # Remove any spaces
    ticket_str = ticket_str.replace(" ", "")

    # Remove 'CW' prefix if present, with or without spaces
    ticket_str = re.sub(r'^(CW\s*-*\s*)', '', ticket_str, flags=re.IGNORECASE)

    # Remove '.X' or '-X' where X is any digit
    ticket_str = re.sub(r'(\.\d+|-\d+)$', '', ticket_str)

    return ticket_str


class GetCWInfo:
    def __init__(self, ticket_id):
        self.CW_BASE_URL = os.getenv("CW_BASE_URL")
        self.CW_COMPANY_ID = os.getenv("CW_COMPANY_ID_PROD")
        self.CW_PUBLIC_KEY = os.getenv('CW_PUBLIC_KEY')
        self.CW_PRIVATE_KEY = os.getenv('CW_PRIVATE_KEY')
        self.ticket_id = normalize_ticket_number(ticket_id)
        self.headers = {"clientid": os.getenv('CW_CLIENT_ID')}
        self.ticket_data = self.get_ticket_by_id()
        self.data = self.get_var()

    @staticmethod
    def process_access_times(custom_fields: List[Dict[str, Any]]) -> Dict[str, str]:
        access_times: Dict[str, Dict[str, str]] = {}
        day_access: Dict[str, str] = {}

        for field in custom_fields:
            value = field.get("value")
            field_name = field.get("caption", "Unknown")

            if "Access Start | " in field_name or "Access End | " in field_name:
                day = field_name.split(" | ")[-1]
                if day not in access_times:
                    access_times[day] = {"start": "00:00", "end": "00:00"}

                if "Start" in field_name:
                    access_times[day]["start"] = value
                elif "End" in field_name:
                    access_times[day]["end"] = value
            else:
                day = field_name.split(" | ")[-1]
                if day in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]:
                    day_access[day] = value

        final_access_times: Dict[str, str] = {}
        for day, times in access_times.items():
            status = day_access.get(day, "Yes").lower()
            start = times["start"]
            end = times["end"]
            if status == "no":
                final_access_times[day] = "No access"
            else:
                final_access_times[day] = f"{start}-{end}" if start != "00:00" and end != "00:00" else "No access"

        return final_access_times

    def get_var(self):
        if not self.ticket_data:
            return {}

        result = {
            "Board": self.get_it("board", "name"),
            "Summary": self.get_it("summary"),
            "Type": self.get_it("type", "name"),
            "Sub-Type": self.get_it("subType", "name"),
            "Status": self.get_it("status", "name"),
            "Customer": self.get_it("company", "name"),
            "Location": f"{self.get_it('city')}, {self.get_it('stateIdentifier')}",
            "Entered by": self.get_it("_info", "enteredBy"),
            "Date entered": self.get_it("_info", "dateEntered"),
            "Products": self.get_ticket_products()
        }

        custom_fields = self.get_it("customFields")
        if custom_fields and isinstance(custom_fields, list):
            result["Access"] = self.process_access_times(custom_fields)
            for field in custom_fields:
                field_name = field.get("caption", "Unknown").strip()
                field_value = field.get("value")
                if field_name and field_value and "Access" not in field_name and all(day not in field_name for day in
                                                                                     ["Monday", "Tuesday", "Wednesday",
                                                                                      "Thursday", "Friday", "Saturday",
                                                                                      "Sunday"]):
                    result[field_name] = field_value

        return {k: v for k, v in result.items() if v}

    def get_it(self, *path):
        data = self.ticket_data
        try:
            for key in path:
                data = data[key]
            return data.strip() if isinstance(data, str) else data
        except (KeyError, TypeError):
            return None

    def get_ticket_by_id(self):
        url = f"{self.CW_BASE_URL}/service/tickets/{self.ticket_id}"
        response = requests.get(url, auth=(f"{self.CW_COMPANY_ID}+{self.CW_PUBLIC_KEY}", self.CW_PRIVATE_KEY),
                                headers=self.headers)

        if response.status_code == 200:
            # print(json.dumps(response.json(), indent=2))
            return response.json()
        elif response.status_code == 404:  # 404 Not Found
            return None  # Return None quietly for "not found" cases
        else:
            print(f"Error fetching ticket {self.ticket_id}: {response.status_code} {response.text}")
            return None

    def get_ticket_products(self):
        products_with_details = {}

        products_url = f"{self.CW_BASE_URL}/procurement/products?conditions=ticket/id={self.ticket_id}"
        response = requests.get(products_url, auth=(f"{self.CW_COMPANY_ID}+{self.CW_PUBLIC_KEY}", self.CW_PRIVATE_KEY),
                                headers=self.headers)
        if response.status_code != 200:
            print(f"Error fetching products for ticket {self.ticket_id}: {response.status_code}")
            return {}

        for product in response.json():
            identifier = product.get("catalogItem", {}).get("identifier")
            if identifier:
                products_with_details[identifier] = {
                    "description": product.get("description"),
                    "quantity": product.get("quantity")
                }

        return products_with_details

    def __str__(self):
        return json.dumps(self.data, indent=2)


class UpdateTicketData:
    def __init__(self, sheet_id=8892937224015748, ticket_column='Equipment Ticket', req_ship_column='Requested Ship'):
        self.sheet_id = sheet_id
        self.escalation_sheet_id = 3968286895067012
        self.ticket_column = ticket_column
        self.req_ship_column = req_ship_column

        self.smart = smartsheet.Smartsheet(os.getenv('SMARTSHEET_ACCESS_TOKEN'))
        self.smart.errors_as_exceptions(True)

        self.sheet = self.fetch_sheet(self.sheet_id)
        self.column_map = {column.title: column.id for column in self.sheet.columns}
        self.escalation_sheet = self.fetch_sheet(self.escalation_sheet_id)
        self.escalation_column_map = {column.title: column.id for column in self.escalation_sheet.columns}

        self.just_len = max(len("* Loaded Smartsheet: " + self.sheet.name), 42)
        logging.warning("*" * (self.just_len + 2))
        logging.warning(("* Loaded Smartsheet: " + self.sheet.name).ljust(self.just_len) + " *")
        logging.warning('* Updating ticket data'.ljust(self.just_len) + " *")
        logging.warning("*" * (self.just_len + 2))
        logging.warning("-" * (self.just_len + 2))
        time.sleep(1)
        # Start the process immediately after initialization
        self.process_rows()

    def extract_ticket_numbers(self, ticket_field):
        # Split by various delimiters to isolate potential ticket numbers
        parts = re.split(r'[\|\,]+', ticket_field)

        ticket_numbers = []
        for part in parts:
            # Remove unnecessary characters and split by spaces to find separate tickets
            for potential_ticket in re.split(r'\s+', part.strip()):
                # Extract numeric parts which might be ticket numbers
                matches = re.findall(r'\d+', potential_ticket)
                ticket_numbers.extend(matches)

        return ticket_numbers

    def fetch_sheet(self, sheet_id):
        sheet = smartsheet_api_call_with_retry(self.smart.Sheets.get_sheet, sheet_id)
        if sheet is None:
            raise Exception("Failed to fetch sheet data after retries")
        return sheet

    def get_value(self, row, column_name, column_map):
        column_id = column_map.get(column_name)
        cell = row.get_column(column_id)

        if cell is None:
            return None

        value = cell.value
        zero_val = isinstance(value, str) and value.startswith("0")

        try:
            value = float(value)
            if value.is_integer():
                value = int(value)
        except (ValueError, TypeError):
            pass

        final_value = "0" + str(value) if zero_val else str(value)
        return final_value

    def pull_date(self, ticket_num):
        try:
            connection = pymssql.connect(server='gp2018', user=f'GRT0\\{os.getenv('GRT_USER')}',
                                         password=os.getenv('GRT_PASS'),
                                         database='SBM01', tds_version="7.0")
        except Error as e:
            logging.warning(str(e).replace("\\n", "\n"))
            sys.exit()
        cursor = connection.cursor()
        cursor.execute(f"""
SELECT DISTINCT 
eq.SOPNUMBE as 'Equipment Ticket', 
eq.CSTPONBR as 'Account Number',
eq.Queue as 'Queue',
COALESCE(CAST(eq.xProject_Name AS NVARCHAR(MAX)), eq.CUSTNAME) as 'Customer Name',
eq.ITEMNMBR as 'Item Number', 
eq.ITEMDESC as 'Item Description',
eq.SERLTNUM as 'Serial Number',
CAST(eq.Notes AS NVARCHAR(MAX)) as 'Internal Notes',
eq.ReqShipDate as 'Requested Ship Date',
eq.CITY as 'City',
eq.STATE as 'State',
eq.Tracking_Number,
eq.USER2ENT as 'SO Creator'
FROM (
SELECT
    sop30300.SOPNUMBE, 
    sop30200.CSTPONBR,
    CASE
        WHEN (sop30200.BACHNUMB IN ('RDY TO INVOICE', 'RDY TO INV') OR sop30200.BACHNUMB LIKE 'Q%') 
        THEN 'RDY TO INVOICE'
        ELSE sop30200.BACHNUMB
    END as Queue,
    sop30200.CUSTNAME,
    spv3SalesDocument.xProject_Name,
    sop30300.ITEMNMBR, 
    sop30300.ITEMDESC,
    sop10201.SERLTNUM,
    spv3SalesDocument.Notes,
    sop30300.ReqShipDate,
    sop30300.CITY,
    sop30300.STATE,
    sop10107.Tracking_Number,
    sop30200.USER2ENT
FROM sop30300
FULL JOIN sop30200 ON sop30200.SOPNUMBE = sop30300.SOPNUMBE
FULL JOIN sop10107 ON sop10107.SOPNUMBE = sop30300.SOPNUMBE
FULL JOIN spv3SalesDocument ON spv3SalesDocument.Sales_Doc_Num = sop30300.SOPNUMBE
FULL JOIN sop10201 ON sop10201.ITEMNMBR = sop30300.ITEMNMBR AND sop10201.SOPNUMBE = sop30300.SOPNUMBE

UNION ALL

SELECT
    sop10100.SOPNUMBE,
    sop10100.CSTPONBR,
    CASE
        WHEN (sop10100.BACHNUMB IN ('RDY TO INVOICE', 'RDY TO INV') OR sop10100.BACHNUMB LIKE 'Q%') 
        THEN 'RDY TO INVOICE'
        ELSE sop10100.BACHNUMB
    END as Queue,
    SOP10100.CUSTNAME,
    spv3SalesDocument.xProject_Name,
    SOP10200.ITEMNMBR, 
    SOP10200.ITEMDESC,
    SOP10201.SERLTNUM,
    spv3SalesDocument.Notes,
    sop10100.ReqShipDate,
    sop10100.CITY,
    sop10100.STATE,
    sop10107.Tracking_Number,
    sop10100.USER2ENT
FROM sop10100
FULL JOIN sop10200 ON sop10200.SOPNUMBE = sop10100.SOPNUMBE
FULL JOIN spv3SalesDocument ON spv3SalesDocument.Sales_Doc_Num = sop10100.SOPNUMBE
FULL JOIN sop10201 ON sop10201.SOPNUMBE = sop10100.SOPNUMBE AND sop10201.ITEMNMBR = sop10200.ITEMNMBR
FULL JOIN sop10107 ON sop10107.SOPNUMBE = sop10100.SOPNUMBE
) AS eq
WHERE eq.SOPNUMBE = '{ticket_num}' OR eq.SOPNUMBE = 'CW{ticket_num}-1'

""")
        return cursor

    def fetch_ticket_owner(self, ticket):
        db_config = {
            "host": "ods",
            "database": "ODS",
            "user": f"GRT0\\{os.getenv('GRT_USER')}",
            "password": os.getenv('GRT_PASS'),
            "tds_version": "7.0"
        }

        api_config = {
            "base_url": "https://api-na.myconnectwise.net/v4_6_release/apis/3.0",
            "company_id": "granitenet",
            "public_key": os.getenv('PUBLIC_KEY'),
            "private_key": os.getenv('PRIVATE_KEY'),
            "headers": {"clientid": os.getenv('CLIENT_ID')}
        }

        # Function to make API request
        def get_ticket_by_id(ticket_id):
            url = f"{api_config['base_url']}/service/tickets/{ticket_id}"
            response = requests.get(url, auth=(
                f"{api_config['company_id']}+{api_config['public_key']}", api_config['private_key']),
                                    headers=api_config['headers'])
            if response.status_code != 200:
                raise Exception
                # return response.json()
            return response.json(), None

        # Connect to the database
        try:
            connection = pymssql.connect(server=db_config['host'], database=db_config['database'],
                                         user=db_config['user'],
                                         password=db_config['password'], tds_version=db_config['tds_version'])
            cursor = connection.cursor()
        except Exception as e:
            logging.error(f"Database connection error: {str(e)}")
            return "Database connection error"

        ticket = ticket.replace("CW", "").replace("-1", "")
        ticket_owner = None
        try:
            cursor.execute(f"""select distinct * from (
            select distinct
            TICKETS_CORE_VIEW.TICKET_ID as 'Ticket',
            TICKETS_CORE_VIEW.MACNUM as 'Account',
            TICKETS_CORE_VIEW.TICKET_TYPE as 'Ticket Type',
            TICKETS_CORE_VIEW.TICKET_SUB_TYPE as 'Ticket Sub-type',
            TICKETS_CORE_VIEW.STATUS as 'Status',
            TICKETS_CORE_VIEW.LOGGED_DT as 'Creation Date',
            TICKETS_CORE_VIEW.SUBJECT as 'Details',
            People.EMPLOYEES.NAME as 'Ticket Creator'

            from Tickets.TICKETS_CORE_VIEW

            full join People.EMPLOYEES on People.EMPLOYEES.EMPLOYEE_ID = TICKETS_CORE_VIEW.LOGGED_BY

            UNION

            select distinct
            Tickets.CAN_TICKETS_CORE.TICKET_ID as 'Ticket',
            Tickets.CAN_TICKETS_CORE.MACNUM as 'Account',
            cast(Tickets.CAN_TICKETS_CORE.TICKET_TYPES_ID as nvarchar) as 'Ticket Type',
            cast(Tickets.CAN_TICKETS_CORE.TICKET_SUB_TYPES_ID as nvarchar) as 'Ticket Sub-type',
            cast(Tickets.CAN_TICKETS_CORE.TICKET_STATUS_ID as nvarchar) as 'Status',
            Tickets.CAN_TICKETS_CORE.LOGGED_DT as 'Creation Date',
            Tickets.CAN_TICKETS_CORE.SUBJECT as 'Details',
            People.EMPLOYEES.NAME as 'Ticket Creator'

            from Tickets.CAN_TICKETS_CORE
            full join People.EMPLOYEES on People.EMPLOYEES.EMPLOYEE_ID = Tickets.CAN_TICKETS_CORE.LOGGED_BY) as MASTER_CS_QUERY

            where [Ticket] = '{ticket}'""")
            ticket_owner = cursor.fetchone()[7]
        except:
            try:
                cursor.execute(f"""select distinct * from (
                select
                WOM.customerinformation$provisioningworkorder.provisioningwonumber as 'WOM Ticket',
                WOM.customerinformation$provisioningworkorder.provwotype as 'WOM Type',
                WOM.customerinformation$provisioningworkorderstatus.[name] as 'WOM Status',
                WOM.customerinformation$provisioningworkorderdetails.createddate as 'WOM Creation Date',
                WOM.usermanagement$account.fullname as 'WOM Created by'


                from 
                WOM.customerinformation$provisioningworkorder

                left join WOM.customerinformation$provisioningworkorderdetails_provisioningworkorder on WOM.customerinformation$provisioningworkorderdetails_provisioningworkorder.customerinformation$provisioningworkorderid = WOM.customerinformation$provisioningworkorder.id
                left join WOM.customerinformation$provisioningworkorderdetails on WOM.customerinformation$provisioningworkorderdetails.id = WOM.customerinformation$provisioningworkorderdetails_provisioningworkorder.customerinformation$provisioningworkorderdetailsid
                left join WOM.usermanagement$account on WOM.usermanagement$account.id = WOM.customerinformation$provisioningworkorder.system$owner
                left join WOM.customerinformation$provisioningworkorder_provisioningworkorderstatus on WOM.customerinformation$provisioningworkorder_provisioningworkorderstatus.customerinformation$provisioningworkorderid = WOM.customerinformation$provisioningworkorder.id
                left join WOM.customerinformation$provisioningworkorderstatus on WOM.customerinformation$provisioningworkorderstatus.id = WOM.customerinformation$provisioningworkorder_provisioningworkorderstatus.customerinformation$provisioningworkorderstatusid	



                ) as MASTER_WOM_QUERY
            where [WOM Ticket] = '{ticket}'""")
                ticket_owner = cursor.fetchone()[4]
            except:
                try:
                    ticket_owner = json.dumps(get_ticket_by_id(ticket)[0]["_info"]["enteredBy"], indent=4).replace(
                        '"', '')
                except:
                    ...
        return ticket_owner

    def process_rows(self):
        errors = []
        for row in self.sheet.rows:
            updated = False
            try:
                tkt = self.get_value(row, self.ticket_column, self.column_map)
                normalized_tkt = normalize_ticket_number(tkt)
                if tkt:
                    sql_row = self.pull_date(tkt).fetchone()
                    if sql_row is not None:
                        sql_date = str(sql_row[8]).strip()
                        sql_stat = str(sql_row[2]).strip()
                        naive_datetime = datetime.datetime.strptime(sql_date, "%Y-%m-%d %H:%M:%S")
                        eastern_timezone = pytz.timezone("America/New_York")
                        add_timezone = naive_datetime.astimezone(eastern_timezone)
                        req_ship_date = add_timezone.strftime("%Y-%m-%d")

                        if req_ship_date != self.get_value(row, self.req_ship_column, self.column_map):
                            new_cell = self.smart.models.Cell()
                            new_cell.column_id = self.column_map[self.req_ship_column]
                            new_cell.value = req_ship_date
                            new_cell.strict = False
                            get_row = smartsheet.models.Row()
                            get_row.id = row.id
                            get_row.cells.append(new_cell)

                            update_call = lambda: self.smart.Sheets.update_rows(self.sheet_id, [get_row])
                            smartsheet_api_call_with_retry(update_call)
                            logging.warning(f"{tkt} - Requested ship date updated: {req_ship_date}")
                            updated = True

                        if sql_stat == "RDY TO INVOICE" and datetime.datetime.now().hour >= 18:
                            new_cell = self.smart.models.Cell()
                            new_cell.column_id = self.column_map["Status"]
                            new_cell.value = "Sent to Shipping"
                            new_cell.strict = False
                            get_row = smartsheet.models.Row()
                            get_row.id = row.id
                            get_row.cells.append(new_cell)

                            update_call = lambda: self.smart.Sheets.update_rows(self.sheet_id, [get_row])
                            smartsheet_api_call_with_retry(update_call)
                            logging.warning(f"{tkt} - Sent to shipping")
                            updated = True

                    originator = self.get_value(row, 'Originator', self.column_map)
                    if originator is None or originator.lower() == "none":
                        originator = self.fetch_ticket_owner(tkt)
                        if originator and 'database' not in originator.lower():
                            new_cell = self.smart.models.Cell()
                            new_cell.column_id = self.column_map['Originator']
                            new_cell.value = originator
                            new_cell.strict = False
                            get_row = smartsheet.models.Row()
                            get_row.id = row.id
                            get_row.cells.append(new_cell)

                            update_call = lambda: self.smart.Sheets.update_rows(self.sheet_id, [get_row])
                            smartsheet_api_call_with_retry(update_call)
                            logging.warning(f"{tkt} - Originator updated: {originator}")
                            updated = True

                    # Flagging logic
                    for escalation_row in self.escalation_sheet.rows:
                        escalated_field = self.get_value(escalation_row, 'Equipment Ticket', self.escalation_column_map)
                        escalated_tickets = self.extract_ticket_numbers(escalated_field)
                        for escalated_tkt in escalated_tickets:
                            normalized_escalated_tkt = normalize_ticket_number(escalated_tkt)

                            if normalized_tkt == normalized_escalated_tkt:
                                is_flagged_value = self.get_value(row, 'Escalated Order', self.column_map)
                                is_flagged = is_flagged_value == '1'
                                # logging.warning(
                                #     f"Ticket {tkt} is {'already flagged' if is_flagged else 'not flagged'} as escalated")
                                if not is_flagged:
                                    # Perform the flag update in the main sheet
                                    new_cell = self.smart.models.Cell()
                                    new_cell.column_id = self.column_map[
                                        'Escalated Order']  # Correct column ID from the main sheet
                                    new_cell.value = 1
                                    new_cell.strict = False

                                    get_row = smartsheet.models.Row()
                                    get_row.id = row.id
                                    get_row.cells.append(new_cell)

                                    update_call = lambda: self.smart.Sheets.update_rows(self.sheet_id, [get_row])
                                    smartsheet_api_call_with_retry(update_call)

                                    logging.warning(f"Flagging ticket {tkt} as escalated")
                                    updated = True
                                    break

                    if not updated:
                        logging.warning("No updates for " + tkt)
                    logging.warning("-" * (self.just_len + 2))

            except Exception as e:
                errors.append(f"Ticket: {tkt} - {e}")

        if errors:
            raise Exception("Errors occurred in the loop: " + '\n'.join(errors))


class SerialColumn:
    def __init__(self, sheet_id=8892937224015748, ticket_column='Equipment Ticket', serial_column='Serial Number(s)',
                 status_column='Status', tkt=None):
        self.sheet_id = sheet_id
        self.ticket_column = ticket_column
        self.serial_column = serial_column
        self.status_column = status_column
        SMARTSHEET_ACCESS_TOKEN = os.getenv('SMARTSHEET_ACCESS_TOKEN')
        smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
        smart.errors_as_exceptions(True)
        # Wrapped call to get the sheet
        sheet = smartsheet_api_call_with_retry(smart.Sheets.get_sheet, sheet_id)
        if sheet is None:
            raise Exception("Failed to fetch sheet data after retries")
        just_len = len("* Loaded Smartsheet: " + sheet.name)
        if just_len < 42:
            just_len = 42
        logging.warning("*" * (just_len + 2))
        logging.warning(("* Loaded Smartsheet: " + sheet.name).ljust(just_len) + " *")
        column_map = {column.title: column.id for column in sheet.columns}

        try:
            connection = pymssql.connect(server='gp2018', user=f'GRT0\\{os.getenv('GRT_USER')}',
                                         password=os.getenv('GRT_PASS'),
                                         database='SBM01')
        except Error as e:
            logging.warning(str(e).replace("\\n", "\n"))
            sys.exit()
        logging.warning('* Connected to SQL Server: GP2018.SBM01'.ljust(just_len) + " *")
        logging.warning('* Updating serial numbers'.ljust(just_len) + " *")
        logging.warning("*" * (just_len + 2))
        time.sleep(1)
        cursor = connection.cursor()

        def get_value(x):
            column_map = {column.title: column.id for column in sheet.columns}
            zero_val = False
            value = row.get_column(column_map[x]).value
            if str(value).startswith("0"):
                zero_val = True
            try:
                value = float(value)
                if value.is_integer():
                    value = int(value)
            except:
                pass
            if zero_val:
                value = "0" + str(value)
            else:
                value = str(value)
            return value

        def is_mac(serial):
            is_hex = True
            for ch in serial:
                if (ch < '0' or ch > '9') and (ch < 'A' or ch > 'F'):
                    is_hex = False

            if is_hex and len(serial) == 12:
                return True
            else:
                return False

        def pull_sns():
            if get_value('Equipment Type') == 'Algo/ATA/Phones':
                cursor.execute(f"""
select distinct dbo.sop10200.sopnumbe, dbo.sop10100.cstponbr, dbo.sop10200.itemnmbr, dbo.sop10201.serltnum, 
dbo.sop10200.itemdesc, BACHNUMB, DBO.IV00101.USCATVLS_1 from dbo.sop10200
full join dbo.sop10100 on dbo.sop10200.sopnumbe = dbo.sop10100.sopnumbe
full join dbo.sop10106 on dbo.sop10200.sopnumbe = dbo.sop10106.sopnumbe
full join dbo.iv00101 on dbo.sop10200.itemnmbr = dbo.iv00101.itemnmbr
full join dbo.iv00102 on dbo.sop10200.itemnmbr = dbo.iv00102.itemnmbr
full join dbo.sop10201 on dbo.sop10200.itemnmbr = dbo.sop10201.itemnmbr and dbo.sop10200.sopnumbe = dbo.sop10201.sopnumbe
where
dbo.sop10200.soptype = '2'
and (dbo.sop10200.sopnumbe like '{tkt}%' or dbo.sop10200.sopnumbe like 'CW{tkt}-1%')
and dbo.sop10201.serltnum not like '%*%'
and sop10200.itemnmbr in ({devices})                   

order by CSTPONBR
            """)
            else:
                cursor.execute(f"""
select distinct dbo.sop10200.sopnumbe, dbo.sop10100.cstponbr, dbo.sop10200.itemnmbr, dbo.sop10201.serltnum,
dbo.sop10200.itemdesc, BACHNUMB, DBO.IV00101.USCATVLS_1
from dbo.sop10200
full join dbo.sop10100 on dbo.sop10200.sopnumbe = dbo.sop10100.sopnumbe
full join dbo.sop10106 on dbo.sop10200.sopnumbe = dbo.sop10106.sopnumbe
full join dbo.iv00101 on dbo.sop10200.itemnmbr = dbo.iv00101.itemnmbr
full join dbo.iv00102 on dbo.sop10200.itemnmbr = dbo.iv00102.itemnmbr
full join dbo.sop10201 on dbo.sop10200.itemnmbr = dbo.sop10201.itemnmbr and dbo.sop10200.sopnumbe = dbo.sop10201.sopnumbe
where
dbo.sop10200.soptype = '2'
and (dbo.sop10200.sopnumbe like '{tkt}%' or dbo.sop10200.sopnumbe like 'CW{tkt}-1%')
and dbo.sop10201.serltnum not like '%*%'
and BACHNUMB in ('CONFIG LAB', 'PROV CONFIG', 'FORTINET CONFIG', 'MOBILITY CONFIG', 'VOIP CONFIG')
and dbo.sop10201.ITEMNMBR not like 'LIC%'
and dbo.sop10200.ITEMDESC not like '%License%'
and dbo.sop10201.serltnum not like '%MAUR0%'
order by CSTPONBR
""")

        for row in sheet.rows:
            mac_flag = False
            if get_value(ticket_column):
                tkt = normalize_ticket_number(get_value(ticket_column))
                pull_sns()
                sql_row = cursor.fetchone()
                sqlRowCount = 0
                sql_list = []
                sn_list = []
                forti_list = []
                item = ""
                while sql_row:
                    sqlRowCount += 1
                    sql_list.append(sqlRowCount)
                    sql_row = cursor.fetchone()
                pull_sns()
                for i in sql_list:
                    sql_row = cursor.fetchone()
                    fortiCheck = sql_row[6].strip()
                    forti_list.append(fortiCheck)
                    if item == sql_row[2].strip():
                        item = sql_row[2].strip()
                        sn = sql_row[3].strip().upper()
                        new_row = f"{sn}"
                    else:
                        item = sql_row[2].strip()
                        sn = sql_row[3].strip().upper()
                        item_pretty = phone_dict[item] if item in phone_dict else item
                        new_row = f"\n[ {item_pretty} ] \n{sn}"
                    if get_value('Equipment Type') == 'Algo/ATA/Phones':
                        if not is_mac(sn):
                            mac_flag = True
                    sn_list.append(new_row)
                sn_list = " \n".join(sn_list).strip().replace("\"", "")
                if sn_list and sn_list != row.get_column(column_map[self.serial_column]).value:
                    # if get_value('MAC Check') == "MAC check complete" or get_value('Config Type (FlexEdge)') in [
                    #     'Configure as new - uCPE', 'Cisco > Flex replacement', 'Flex > Flex replacement']:
                    if get_value('MAC Check') == "MAC check complete":
                        pass
                    else:
                        logging.warning(f"{tkt}\n")
                        logging.warning(sn_list)

                        # Add SNs to Serial Number column
                        new_cell = smart.models.Cell()
                        new_cell.column_id = column_map[serial_column]
                        new_cell.value = sn_list
                        new_cell.strict = False
                        get_row = smart.models.Row()
                        get_row.id = row.id
                        get_row.cells.append(new_cell)
                        # Wrap the update_rows call in a lambda to defer execution
                        update_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

                        # Use the retry function to attempt the update
                        updated_row = smartsheet_api_call_with_retry(update_call)
                        if mac_flag:
                            new_cell = smart.models.Cell()
                            new_cell.column_id = column_map['MAC Check']
                            new_cell.value = "CHECK MAC ADDRESSES"
                            new_cell.strict = False
                            get_row = smart.models.Row()
                            get_row.id = row.id
                            get_row.cells.append(new_cell)
                            # Wrap the update_rows call in a lambda to defer execution
                            update_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

                            # Use the retry function to attempt the update
                            updated_row = smartsheet_api_call_with_retry(update_call)

                        # Update Status column
                        if row.get_column(column_map[status_column]).value in ["Allocated", "Pending Allocation",
                                                                               "Unworked", "Equipment ON ORDER",
                                                                               "Pending SO"]:
                            new_cell = smart.models.Cell()
                            new_cell.column_id = column_map[status_column]
                            new_cell.value = "Allocated"
                            new_cell.strict = False
                            get_row = smart.models.Row()
                            get_row.id = row.id
                            get_row.cells.append(new_cell)
                            # Wrap the update_rows call in a lambda to defer execution
                            update_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

                            # Use the retry function to attempt the update
                            updated_row = smartsheet_api_call_with_retry(update_call)

                elif not sn_list:
                    logging.warning(f"No equipment allocated for {tkt}")
                else:
                    logging.warning(f"Serial(s) already entered for {tkt}")

                logging.warning("-" * (just_len + 2))


class DIAAutoTemplate:
    def __init__(self):
        geolocator = Nominatim(user_agent='geoapiExercises')
        SMARTSHEET_ACCESS_TOKEN = os.getenv('SMARTSHEET_ACCESS_TOKEN')
        smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
        smart.errors_as_exceptions(True)
        sheet_id = 8892937224015748
        # Wrapped call to get the sheet
        sheet = smartsheet_api_call_with_retry(smart.Sheets.get_sheet, sheet_id)
        if sheet is None:
            raise Exception("Failed to fetch sheet data after retries")
        just_len = len("* Loaded Smartsheet: " + sheet.name)
        if just_len < 42:
            just_len = 42
        column_map = {column.title: column.id for column in sheet.columns}
        logging.warning("*" * (just_len + 2))
        logging.warning(("* Loaded Smartsheet: " + sheet.name).ljust(just_len) + " *")
        logging.warning(("* Compiling templates...".ljust(just_len) + " *"))
        logging.warning("*" * (just_len + 2))

        def get_value(x):
            zero_val = False
            value = row.get_column(column_map[x]).value
            if str(value).startswith("0"):
                zero_val = True
            try:
                value = float(value)
                if value.is_integer():
                    value = int(value)
            except:
                pass
            if zero_val:
                value = "0" + str(value)
            else:
                value = str(value)
            return value

        def update_cell(smart_column, smart_value):
            new_cell = smart.models.Cell()
            new_cell.column_id = column_map[smart_column]
            new_cell.value = smart_value
            new_cell.strict = False
            get_row = smart.models.Row()
            get_row.id = row.id
            get_row.cells.append(new_cell)
            response = smart.Sheets.update_rows(sheet_id, [get_row])

        class Attachments:
            def __init__(self, smartsheet_obj):
                """Init Attachments with base Smartsheet object."""
                self._base = smartsheet_obj
                self._log = logging.getLogger(__name__)

            def attach_file_to_row(self, sheet_id, row_id, _file):
                if not all(val is not None for val in ['sheet_id', 'row_id', '_file']):
                    raise ValueError(
                        ('One or more required values '
                         'are missing from call to ' + __name__))

                _op = fresh_operation('attach_file_to_row')
                _op['method'] = 'POST'
                _op['path'] = '/sheets/' + str(sheet_id) + '/rows/' + str(
                    row_id) + '/attachments'
                _op['files'] = {}
                _op['files']['file'] = _file
                _op['filename'] = _file

                expected = ['Result', 'Attachment']

                prepped_request = self._base.prepare_request(_op)
                response = self._base.request(prepped_request, expected, _op)

                return response

        def attach_file():
            newfile = f'{hostname}.txt'
            if json.loads(str(smart.Search.search_sheet(sheet_id=8892937224015748, query=newfile)))["totalCount"] == 0:
                with open(newfile, 'w') as f:
                    f.write(write)
                updated_attachment = smart.Attachments.attach_file_to_row(
                    sheet_id,  # sheet_id
                    row_id,  # row_id
                    (newfile,
                     open(newfile, 'rb'),
                     'application/notepad')
                )
                f.close()
                os.remove(newfile)
                logging.warning(f"{newfile} written!")
                update_cell('Template Status', 'Template attached')


            else:
                logging.warning(f"{newfile} already exists!")
            logging.warning("-" * (just_len + 2))

        for row in sheet.rows:
            row_id = row.id
            if get_value('Config Type (Cisco)') == 'Config lab (template)':
                tkt = normalize_ticket_number(get_value('Equipment Ticket'))
                status = get_value('Status')
                if status in ('Allocated', 'In Process', 'None', 'Pending Information'):
                    account = get_value('Child Account')
                    configType = get_value('DIA/T1 Config Template')
                    hostname = get_value('Hostname')
                    lanSubnet = get_value('LAN Subnet Mask')
                    lanGateway = get_value('LAN Gateway IP')
                    lanNetwork = get_value('LAN Network IP')
                    cktId = get_value('Circuit ID')
                    wanSubnet = get_value('WAN Subnet Mask')
                    wanGateway = get_value('WAN Gateway IP')
                    lanIp = get_value('LAN Usable IP')
                    wanIp = get_value('WAN Usable IP')
                    carrier = get_value('Carrier')
                    loopback66 = get_value('Loopback IP')
                    loopback66 = re.sub('/..', '', loopback66)
                    speed = str(get_value('Speed'))
                    speed = re.sub('\\D', '', speed)
                    address = str(get_value("City, State/Province"))
                    address = usaddress.tag(address)
                    address = json.loads(json.dumps(address[0]))
                    encap_vlan = get_value('VLAN')
                    encap_in_vlan = get_value('Inner VLAN')
                    encap_out_vlan = get_value('Outer VLAN')
                    try:
                        password = address['PlaceName'].replace(',', '').replace(' ', '') + speed
                    except:
                        password = str(get_value("City, State/Province").replace(',', '').replace(' ', '')) + speed
                    try:
                        city = address['PlaceName'].replace(',', '')
                        state = address['StateName']
                        location = f"{city}, {state}"
                    except:
                        location = str(get_value("City, State/Province"))
                    if configType == 'On-net_ASR-920':
                        varCheck = {"Equipment Ticket": tkt,
                                    "Child Account": account,
                                    "DIA/T1 Config Template": configType,
                                    "Hostname": hostname,
                                    "LAN Subnet Mask": lanSubnet,
                                    "LAN Gateway IP": lanGateway,
                                    "Circuit ID": cktId,
                                    "WAN Subnet Mask": wanSubnet,
                                    "WAN Usable IP": wanIp,
                                    "WAN Gateway IP": wanGateway,
                                    "Loopback IP": loopback66,
                                    "Carrier": carrier,
                                    "City, State/Province": address}
                        if any([i in carrier.lower() for i in ("att", "at&t")]):
                            varCheck["Inner VLAN"] = encap_in_vlan
                            varCheck["Outer VLAN"] = encap_out_vlan
                            encapsulation = f"""encapsulation dot1q {encap_out_vlan} second-dot1q {encap_in_vlan}
    rewrite ingress tag pop 2 symmetric"""
                        elif any([i in carrier.lower() for i in ("fairpoint", "consolidated", "frontier", "ziply")]):
                            varCheck["Encapsulation VLAN"] = encap_vlan
                            encapsulation = f"""encapsulation dot1q {encap_vlan}
    rewrite ingress tag pop 1 symmetric"""
                        else:
                            encapsulation = "encapsulation untagged"

                    else:
                        varCheck = {"Equipment Ticket": tkt,
                                    "Child Account": account,
                                    "DIA/T1 Config Template": configType,
                                    "Hostname": hostname,
                                    "LAN Subnet Mask": lanSubnet,
                                    "LAN Gateway IP": lanGateway,
                                    "LAN Network IP": lanNetwork,
                                    "Circuit ID": cktId,
                                    "WAN Subnet Mask": wanSubnet,
                                    "WAN Gateway IP": wanGateway,
                                    "LAN Usable IP": lanIp,
                                    "WAN Usable IP": wanIp,
                                    "Speed": speed,
                                    "City, State/Province": address}
                    varList = []
                    missingData = False
                    for var in varCheck:
                        if varCheck[var] in ["None", "none", "n/a", "na", "N/A", "NA"] or varCheck[var] is None:
                            missingData = True
                            varList.append(var)
                    if missingData:
                        logging.warning(
                            textwrap.fill(f'{tkt} missing data in column(s): {varList}', width=(just_len + 2)))
                        var_notes = str(varList).replace("[", "").replace("]", "").replace("'", "")
                        missing_notes = f'\nMissing data: {var_notes}'
                        update_cell('Notes', missing_notes)
                        update_cell('Template Status', 'Missing data')
                        update_cell('Status', 'Pending Information')
                    else:
                        config_present = True
                        if configType == '4K series RJ45':
                            write = f"""hostname {hostname}


    aaa new-model

    ip tacacs source-interface GigabitEthernet0/0/0

    aaa authentication login default group tacacs+ local
    aaa authentication enable default group tacacs+ enable
    aaa authorization console
    aaa authorization exec default group tacacs+ local 
    aaa authorization commands 1 default group tacacs+ local 
    aaa authorization commands 15 default group tacacs+ local 
    aaa authorization network default group tacacs+ local 
    aaa accounting exec default start-stop group tacacs+
    aaa accounting commands 1 default start-stop group tacacs+
    aaa accounting commands 15 default start-stop group tacacs+
    aaa accounting network default start-stop group tacacs+
    aaa accounting connection default start-stop group tacacs+
    aaa accounting system default start-stop group tacacs+

    tacacs server EXTERNAL_ACS
    address ipv4 172.85.135.235 
    timeout 1 
    key V#38fe;5K[

    tacacs-server directed-request
    !
    !
    !
    enable secret Granite1!
    username GraniteNOC secret Gran1te0ff
    username ADSOffnet secret {password}
    username TempTech secret Password123

    service password-encryption



    ip domain name granitenet.com


    crypto key generate rsa modulus 2048

    ip ssh version 2
    banner motd #
    WARNING: To protect the system from unauthorized use and to ensure
    that the system is functioning properly, activities on this system are
    monitored and recorded and subject to audit.  Use of this system is
    expressed consent to such monitoring and recording.  Any unauthorized
    access or use of this Automated Information System is prohibited and
    could be subject to criminal and civil penalties.
    #

    !
    ip domain-lookup
    ip name-server 8.8.8.8 4.2.2.2
    service call-home
    call-home
    contact-email-addr sch-smart-licensing@cisco.com
    profile "CiscoTAC-1"
    active
    destination transport-method http
    exit
    license smart enable
    !

    interface GigabitEthernet0/0/0
    description {cktId} // {account}
    ip address {wanIp} {wanSubnet}
    media-type  RJ45
    no shutdown
    exit


    interface GigabitEthernet0/0/1
    description To LAN
    ip address {lanGateway} {lanSubnet}
    negotiation auto
    no shutdown

    ip route 0.0.0.0 0.0.0.0 {wanGateway}
    ip ssh version 2


    line con 0
    login local
    no ip access-list standard CPEAccess
    ip access-list standard CPEAccess
    permit 172.16.0.0 0.15.255.255
    permit 10.0.0.0 0.255.255.255
    permit 192.168.0.0 0.0.255.255
    permit 198.18.0.0 0.1.255.255
    permit 198.51.100.0 0.0.0.255
    permit host 65.202.145.2
    permit host 72.46.171.2
    permit host 172.85.135.238
    permit host 162.223.83.42
    permit host 162.223.83.38
    deny   any

    ip access-list standard GRANITESNMP
    permit 162.223.83.38
    permit 162.223.86.38
    permit 172.85.135.238
    permit 198.19.0.33
    permit 198.19.0.32
    deny   any log
    !

    SNMP-Server view SNMPv3View 1.3.6 included
    SNMP-Server group SNMPv3Group v3 priv Read SNMPv3View Write SNMPv3View
    SNMP-Server user SNMPv3User SNMPv3Group v3 auth md5 $$w1n$t@R!$$ priv des $$w1n$t@R!$$ access GRANITESNMP

    archive
    log config
    logging enable
    logging size 150
    notify syslog contenttype plaintext
    hidekeys

    cts logging verbose

    clock timezone EDT -5 0
    clock summer-time EDT recurring 2 Sun Mar 3:00 1 Sun Nov 3:00

    line con 0
    exec-timeout 15 0
    logging synchronous
    stopbits 1
    line vty 0 4
    access-class CPEAccess in
    exec-timeout 15 0
    logging synchronous
    transport input ssh
    line vty 5 15
    access-class CPEAccess in
    exec-timeout 5 0
    logging synchronous
    transport input ssh

    no ip nat service sip udp port 5060
    ip forward-protocol nd
    no ip http server
    no ip http secure-server
    !
    !
    !

    line con 0
    exec-timeout 15 0
    logging synchronous
    line vty 0 15
    logging synchronous
    transport input ssh
    end


    copy run start
    copy run start

    sh ver

    sh run
                        """

                        elif configType == '4K series with SFP':
                            write = f"""hostname {hostname}

    aaa new-model

    ip tacacs source-interface GigabitEthernet0/0/0

    aaa authentication login default group tacacs+ local
    aaa authentication enable default group tacacs+ enable
    aaa authorization console
    aaa authorization exec default group tacacs+ local 
    aaa authorization commands 1 default group tacacs+ local 
    aaa authorization commands 15 default group tacacs+ local 
    aaa authorization network default group tacacs+ local 
    aaa accounting exec default start-stop group tacacs+
    aaa accounting commands 1 default start-stop group tacacs+
    aaa accounting commands 15 default start-stop group tacacs+
    aaa accounting network default start-stop group tacacs+
    aaa accounting connection default start-stop group tacacs+
    aaa accounting system default start-stop group tacacs+

    tacacs server EXTERNAL_ACS
    address ipv4 172.85.135.235 
    timeout 1 
    key V#38fe;5K[

    tacacs-server directed-request

    !
    !
    !                        

    enable secret Granite1!
    username GraniteNOC secret Gran1te0ff
    username ADSOffnet secret {password}
    username TempTech secret Password123


    service password-encryption


    ip domain name granitenet.com


    crypto key generate rsa modulus 2048




    ip ssh version 2
    banner motd #
    WARNING: To protect the system from unauthorized use and to ensure
    that the system is functioning properly, activities on this system are
    monitored and recorded and subject to audit.  Use of this system is
    expressed consent to such monitoring and recording.  Any unauthorized
    access or use of this Automated Information System is prohibited and
    could be subject to criminal and civil penalties.
    #

    !
    ip domain-lookup
    ip name-server 8.8.8.8 4.2.2.2
    service call-home
    call-home
    contact-email-addr sch-smart-licensing@cisco.com
    profile "CiscoTAC-1"
    active
    destination transport-method http
    exit
    license smart enable
    !


    interface GigabitEthernet0/0/0
    description {cktId} // {account}
    ip address {wanIp} {wanSubnet}
    media-type  sfp
    no shutdown
    exit


    interface GigabitEthernet0/0/1
    description To LAN
    ip address {lanGateway} {lanSubnet}
    negotiation auto
    no shutdown

    ip route 0.0.0.0 0.0.0.0 {wanGateway}
    ip ssh version 2


    line con 0
    login local
    no ip access-list standard CPEAccess
    ip access-list standard CPEAccess
    permit 172.16.0.0 0.15.255.255
    permit 10.0.0.0 0.255.255.255
    permit 192.168.0.0 0.0.255.255
    permit 198.18.0.0 0.1.255.255
    permit 198.51.100.0 0.0.0.255
    permit host 65.202.145.2
    permit host 72.46.171.2
    permit host 172.85.135.238
    permit host 162.223.83.42
    permit host 162.223.83.38
    deny   any
    !
    ip access-list standard GRANITESNMP
    permit 162.223.83.38
    permit 162.223.86.38
    permit 172.85.135.238
    permit 198.19.0.33
    permit 198.19.0.32
    deny   any log
    !
    !
    SNMP-Server view SNMPv3View 1.3.6 included
    SNMP-Server group SNMPv3Group v3 priv Read SNMPv3View Write SNMPv3View
    SNMP-Server user SNMPv3User SNMPv3Group v3 auth md5 $$w1n$t@R!$$ priv des $$w1n$t@R!$$ access GRANITESNMP
    !
    archive
    log config
    logging enable
    logging size 150
    notify syslog contenttype plaintext
    hidekeys
    !
    cts logging verbose

    clock timezone EDT -5 0
    clock summer-time EDT recurring 2 Sun Mar 3:00 1 Sun Nov 3:00

    line con 0
    exec-timeout 15 0
    logging synchronous
    stopbits 1
    line vty 0 4
    access-class CPEAccess in
    exec-timeout 15 0
    logging synchronous
    transport input ssh
    line vty 5 15
    access-class CPEAccess in
    exec-timeout 5 0
    logging synchronous
    transport input ssh

    no ip nat service sip udp port 5060
    ip forward-protocol nd
    no ip http server
    no ip http secure-server
    !
    !
    !
    !
    line con 0
    exec-timeout 15 0
    logging synchronous
    line vty 0 15
    logging synchronous
    transport input ssh
    end
    !
                        """

                        elif configType == 'Offnet_ASR-920':
                            write = f"""!
    no service pad
    service timestamps debug datetime msec
    service timestamps log datetime msec
    service password-encryption
    no platform punt-keepalive disable-kernel-core
    platform bfd-debug-trace 1
    platform xconnect load-balance-hash-algo mac-ip-instanceid
    platform tcam-parity-error enable
    platform tcam-threshold alarm-frequency 1
    !
    hostname {hostname}
    !
    boot-start-marker
    boot-end-marker
    !
    !
    vrf definition Mgmt-intf
    !
    address-family ipv4
    exit-address-family
    !
    address-family ipv6
    exit-address-family
    !
    logging buffered 51200 warnings
    !
    aaa new-model
    !
    !
    aaa group server tacacs+ management
    server-private 172.85.135.235 timeout 1 key 7 122F46444A0D095F7F001F
    ip tacacs source-interface BDI100
    !
    aaa authentication login default group management local
    aaa authentication enable default group management enable
    aaa authorization console
    aaa authorization exec default group management local 
    aaa authorization commands 1 default group management local 
    aaa authorization commands 15 default group management local 
    aaa authorization network default group management local 
    aaa accounting exec default start-stop group management
    aaa accounting commands 1 default start-stop group management
    aaa accounting commands 15 default start-stop group management
    aaa accounting network default start-stop group management
    aaa accounting connection default start-stop group management
    aaa accounting system default start-stop group management
    !
    !
    !
    !
    !
    aaa session-id common
    clock timezone EST -5 0
    clock summer-time EDT recurring 2 Sun Mar 3:00 1 Sun Nov 3:00
    facility-alarm critical exceed-action shutdown
    !
    !
    !
    !
    !
    !
    !
    !
    !

    no ip domain lookup
    ip domain name granitempls.com
    !
    ip dhcp pool Mgmt-intf
    network 172.16.0.0 255.255.255.0
    !
    !
    !
    !
    login block-for 300 attempts 4 within 120
    login delay 2
    login on-failure log
    login on-success log
    !
    !
    !         
    !
    !
    !
    !
    !
    !
    multilink bundle-name authenticated
    !
    !
    !
    sdm prefer default 
    !
    username NOCAdmin privilege 15 secret 5 $1$QxBT$i0o24EYUorGW8MGdm8.gE1
    username turnup-temp privilege 15 secret 5 $1$Mk4X$0a44f6jFsUke7lwozYCt5/
    username granitenoc secret 5 $1$J.jq$fKW6pxUQp.gCyuCxGz/lf0
    !
    redundancy
    bridge-domain 100 
    !
    !
    !         
    !
    !
    transceiver type all
    monitoring
    !
    ! 
    !
    crypto key generate rsa modulus 2048
    !
    !
    !
    !
    ip domain-lookup
    ip name-server 8.8.8.8 4.2.2.2
    service call-home
    call-home
    contact-email-addr sch-smart-licensing@cisco.com
    profile "CiscoTAC-1"
    active
    destination transport-method http
    exit
    license smart enable
    !
    !
    ---------------------------------------------------------------------
    -----------------------*** BREAK COPY HERE ***-----------------------
    ---------------------------------------------------------------------
    !
    !
    interface GigabitEthernet0/0/0
    description  {cktId} {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface GigabitEthernet0/0/1
    description CUSTOMER LAN
    no ip address
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    negotiation auto
    !
    !
    interface GigabitEthernet0/0/4
    description  {cktId} {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface GigabitEthernet0/0/5
    description CUSTOMER LAN
    no ip address
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    negotiation auto
    !
    interface TenGigabitEthernet0/0/2
    description {cktId} // {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface TenGigabitEthernet0/0/3
    description OPTICAL CUSTOMER LAN
    no ip address
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    !
    !
    interface TenGigabitEthernet0/0/12
    description {cktId} // {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface TenGigabitEthernet0/0/13
    description OPTICAL CUSTOMER LAN
    no ip address
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    !
    !
    interface TenGigabitEthernet0/0/24
    description {cktId} // {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface TenGigabitEthernet0/0/25
    description OPTICAL CUSTOMER LAN
    no ip address
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    !
    interface TenGigabitEthernet0/0/4
    no ip address
    !
    interface TenGigabitEthernet0/0/5
    no ip address
    !
    interface GigabitEthernet0
    vrf forwarding Mgmt-intf
    ip address 172.16.0.1 255.255.255.0
    negotiation auto
    !
    interface BDI100
    ip address  {wanIp} {wanSubnet}
    no shut
    !
    interface BDI200
    ip address {lanGateway} {lanSubnet}
    no shutdown
    !
    ip forward-protocol nd
    !
    ip bgp-community new-format
    no ip http server
    no ip http secure-server
    ip tftp source-interface GigabitEthernet0
    ip tacacs source-interface BDI100
    ip ssh time-out 60
    ip ssh authentication-retries 2
    ip ssh version 2
    ip ssh server algorithm encryption aes128-ctr aes192-ctr aes256-ctr aes128-cbc aes192-cbc aes256-cbc 3des-cbc
    ip ssh client algorithm encryption aes128-ctr aes192-ctr aes256-ctr aes128-cbc aes192-cbc aes256-cbc 3des-cbc
    ip route 0.0.0.0 0.0.0.0 {wanGateway}
    !
    ip access-list standard CPEAccess
    permit 162.223.83.42
    permit 162.223.83.38
    permit 65.202.145.2
    permit 134.204.10.2
    permit 134.204.11.2
    permit 172.85.135.238
    permit 172.85.180.240
    permit 72.46.171.2
    permit 172.16.0.0 0.15.255.255
    permit 10.0.0.0 0.255.255.255
    permit 192.168.0.0 0.0.255.255
    permit 198.18.0.0 0.1.255.255
    permit 198.51.100.0 0.0.0.255
    permit 100.64.1.0 0.0.0.255
    permit 100.64.2.0 0.0.0.255
    deny   any
    ip access-list standard GRANITESNMP
    permit 162.223.83.42
    permit 162.223.83.38
    permit 162.223.86.38
    permit 172.85.135.238
    permit 172.85.180.240
    permit 198.19.0.33
    permit 198.19.0.32
    permit 198.19.127.0 0.0.0.31
    deny   any log
    ip access-list standard GraniteNTP
    permit 162.223.83.36
    permit 162.223.86.36
    permit 198.18.0.0 0.1.255.255
    deny   any
    !
    !
    !
    snmp-server group SNMPv3Group v3 priv read SNMPv3View write SNMPv3View 
    snmp-server view SNMPv3View dod included
    SNMP-Server user SNMPv3User SNMPv3Group v3 auth md5 $$w1n$t@R!$$ priv des $$w1n$t@R!$$ access GRANITESNMP
    snmp-server community yNLQ14xxH4mgV RO GRANITESNMP
    snmp-server trap-source BDI100
    snmp-server location {location}
    snmp-server chassis-id {hostname}
    snmp-server enable traps bfd
    snmp-server enable traps config-copy
    snmp-server enable traps config
    snmp-server enable traps event-manager
    snmp-server enable traps cpu threshold
    snmp-server enable traps ethernet evc status create delete
    snmp-server enable traps alarms informational
    snmp-server enable traps ethernet cfm alarm
    snmp-server enable traps transceiver all

    snmp ifmib ifindex persist
    !
    tacacs-server directed-request
    !
    !
    !
    control-plane
    !
    banner motd ^CCCCC

    WARNING: To protect the system from unauthorized use and to ensure
    that the system is functioning properly, activities on this system are
    monitored and recorded and subject to audit.  Use of this system is
    expressed consent to such monitoring and recording.  Any unauthorized
    access or use of this Automated Information System is prohibited and
    could be subject to criminal and civil penalties.

    ^C
    !
    line con 0
    exec-timeout 15 0
    logging synchronous
    stopbits 1
    line vty 0 4
    access-class CPEAccess in vrf-also
    logging synchronous
    transport input ssh
    line vty 5 15
    access-class CPEAccess in vrf-also
    logging synchronous
    transport input ssh
    !         
    exception crashinfo file bootflash:crashinfo
    ntp source BDI100
    ntp access-group peer GraniteNTP
    ntp server 162.223.83.36 prefer
    ntp server 162.223.86.36
    !
    !
    end
                        """

                        elif configType == 'Offnet_ASR-920_24Port':
                            write = f"""hostname {hostname}

    boot-start-marker	
    boot-end-marker


    vrf definition Mgmt-intf

    address-family ipv4
    exit-address-family

    address-family ipv6
    exit-address-family

    logging buffered 51200 warnings

    aaa new-model


    aaa group server tacacs+ management
    server-private 172.85.135.235 timeout 1 key 7 122F46444A0D095F7F001F
    ip tacacs source-interface BDI100

    aaa authentication login default group management local
    aaa authentication enable default group management enable
    aaa authorization console
    aaa authorization exec default group management local 
    aaa authorization commands 1 default group management local 
    aaa authorization commands 15 default group management local 
    aaa authorization network default group management local 
    aaa accounting exec default start-stop group management
    aaa accounting commands 1 default start-stop group management
    aaa accounting commands 15 default start-stop group management
    aaa accounting network default start-stop group management
    aaa accounting connection default start-stop group management
    aaa accounting system default start-stop group management





    aaa session-id common
    facility-alarm critical exceed-action shutdown











    no ip domain lookup
    ip domain name granitempls.com












    multilink bundle-name authenticated


    license boot level metroipaccess

    sdm prefer default 



    username NOCAdmin privilege 15 secret 5 $1$QxBT$i0o24EYUorGW8MGdm8.gE1
    username turnup-temp privilege 15 secret 5 $1$Mk4X$0a44f6jFsUke7lwozYCt5/
    username granitenoc secret 5 $1$J.jq$fKW6pxUQp.gCyuCxGz/lf0

    redundancy
    bridge-domain 100 





    transceiver type all
    monitoring







    crypto key generate rsa modulus 2048	  


    ip dhcp pool Mgmt-intf
    network 172.16.0.0 255.255.255.0          


    !
    ip domain-lookup
    ip name-server 8.8.8.8 4.2.2.2
    service call-home
    call-home
    contact-email-addr sch-smart-licensing@cisco.com
    profile "CiscoTAC-1"
    active
    destination transport-method http
    exit
    license smart enable
    !



    interface GigabitEthernet0/0/0
    description {cktId} // {account}
    no ip address
    media-type rj45
    negotiation auto
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    no shutdown



    interface GigabitEthernet0/0/1
    description // CUSTOMER_PUBLIC_LAN //
    media-type rj45
    negotiation auto
    no shutdown

    interface range GigabitEthernet0/0/2-23
    shut

    interface TenGigabitEthernet0/0/24
    description {cktId} // {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    no shutdown


    interface TenGigabitEthernet0/0/25
    description // CUSTOMER_PUBLIC_LAN_Fiber//
    ip address  {lanIp} {lanSubnet}
    negotiation auto
    no shutdown

    interface range TenGigabitEthernet0/0/26-27
    shut


    interface GigabitEthernet0
    vrf forwarding Mgmt-intf
    no ip address
    negotiation auto

    interface BDI100
    ip address  {wanIp} {wanSubnet}

    no shutdown

    ip forward-protocol nd

    ip bgp-community new-format
    no ip http server
    no ip http secure-server
    ip tftp source-interface GigabitEthernet0
    ip ssh time-out 60
    ip ssh authentication-retries 2
    ip ssh version 2
    ip ssh server algorithm encryption aes128-ctr aes192-ctr aes256-ctr aes128-cbc aes192-cbc aes256-cbc 3des-cbc
    ip ssh client algorithm encryption aes128-ctr aes192-ctr aes256-ctr aes128-cbc aes192-cbc aes256-cbc 3des-cbc
    ip route 0.0.0.0 0.0.0.0 {wanGateway}



    ip access-list standard GRANITESNMP
    permit 162.223.83.42
    permit 162.223.83.38
    permit 162.223.86.38
    permit 172.85.135.238
    permit 172.85.180.240
    permit 198.19.0.33
    permit 198.19.0.32
    permit 198.19.127.0 0.0.0.31
    deny   any log
    ip access-list standard GraniteNTP
    permit 162.223.83.36
    permit 162.223.86.36
    permit 198.18.0.0 0.1.255.255
    deny   any


    ip access-list extended VTY
    permit ip host 162.223.83.42 any
    permit ip host 162.223.83.38 any
    permit ip host 65.202.145.2 any
    permit ip host 172.85.135.238 any
    permit ip host 198.19.127.2 any
    permit ip host 198.19.0.21 any
    permit ip host 172.85.228.254 any
    permit ip host 134.204.10.2 any
    permit ip host 134.204.11.2 any
    permit ip host 72.46.171.2 any
    permit ip 172.16.0.0 0.15.255.255 any
    permit ip 10.0.0.0 0.255.255.255 any
    permit ip 192.168.0.0 0.0.255.255 any
    permit ip 198.18.0.0 0.1.255.255 any
    permit ip 100.64.1.0 0.0.0.255 any
    permit ip 198.51.100.0 0.0.0.255 any
    permit ip 100.64.2.0 0.0.0.255 any


    snmp-server group SNMPv3Group v3 priv read SNMPv3View write SNMPv3View 
    snmp-server view SNMPv3View dod included
    SNMP-Server user SNMPv3User SNMPv3Group v3 auth md5 $$w1n$t@R $$ priv des $$w1n$t@R $$ access GRANITESNMP
    snmp-server trap-source BDI100
    snmp-server location {location}
    snmp-server chassis-id {hostname}
    snmp-server enable traps bfd
    snmp-server enable traps config-copy
    snmp-server enable traps config
    snmp-server enable traps event-manager
    snmp-server enable traps cpu threshold
    snmp-server enable traps ethernet evc status create delete
    snmp-server enable traps alarms informational
    snmp-server enable traps ethernet cfm alarm
    snmp-server enable traps transceiver all
    snmp ifmib ifindex persist

    tacacs-server directed-request





    control-plane

    banner motd ^CCCCCC

    WARNING: To protect the system from unauthorized use and to ensure
    that the system is functioning properly, activities on this system are
    monitored and recorded and subject to audit.  Use of this system is
    expressed consent to such monitoring and recording.  Any unauthorized
    access or use of this Automated Information System is prohibited and
    could be subject to criminal and civil penalties.

    ^C

    line con 0
    exec-timeout 15 0
    logging synchronous
    stopbits 1
    line vty 0 4
    access-class VTY in vrf-also
    logging synchronous
    transport input ssh
    line vty 5 15
    access-class VTY in vrf-also
    logging synchronous
    transport input ssh

    exception crashinfo file bootflash:crashinfo

    ntp source BDI100
    ntp access-group peer GraniteNTP
    ntp server 162.223.83.36 prefer
    ntp server 162.223.86.36



    end
                        """

                        elif configType == 'Offnet_ASR-920_DHCP_LAN_POOL':
                            write = f"""!
    no service pad
    service timestamps debug datetime msec
    service timestamps log datetime msec
    service password-encryption
    no platform punt-keepalive disable-kernel-core
    platform bfd-debug-trace 1
    platform xconnect load-balance-hash-algo mac-ip-instanceid
    platform tcam-parity-error enable
    platform tcam-threshold alarm-frequency 1
    !
    hostname {hostname}
    !
    boot-start-marker
    boot-end-marker
    !
    !
    vrf definition Mgmt-intf
    !
    address-family ipv4
    exit-address-family
    !
    address-family ipv6
    exit-address-family
    !
    logging buffered 51200 warnings
    !
    aaa new-model
    !
    !
    aaa group server tacacs+ management
    server-private 172.85.135.235 timeout 1 key 7 122F46444A0D095F7F001F
    ip tacacs source-interface BDI100
    !
    aaa authentication login default group management local
    aaa authentication enable default group management enable
    aaa authorization console
    aaa authorization exec default group management local 
    aaa authorization commands 1 default group management local 
    aaa authorization commands 15 default group management local 
    aaa authorization network default group management local 
    aaa accounting exec default start-stop group management
    aaa accounting commands 1 default start-stop group management
    aaa accounting commands 15 default start-stop group management
    aaa accounting network default start-stop group management
    aaa accounting connection default start-stop group management
    aaa accounting system default start-stop group management
    !
    !
    !
    !
    !
    aaa session-id common
    clock timezone EST -5 0
    clock summer-time EDT recurring 2 Sun Mar 3:00 1 Sun Nov 3:00
    facility-alarm critical exceed-action shutdown
    !
    !
    !
    !
    !
    !
    !
    !
    !

    no ip domain lookup
    ip domain name granitempls.com
    ip dhcp excluded-address {lanGateway}
    !
    ip dhcp pool Mgmt-intf
    network 172.16.0.0 255.255.255.0
    !
    ip dhcp pool PUBLIC
    network {lanIp} {lanSubnet}
    default-router {lanGateway}
    dns-server 8.8.8.8 8.8.4.4 
    !
    !
    !
    login block-for 300 attempts 4 within 120
    login delay 2
    login on-failure log
    login on-success log
    !
    !
    !         
    !
    !
    !
    !
    !
    !
    multilink bundle-name authenticated
    !
    !
    !
    sdm prefer default 
    !
    username NOCAdmin privilege 15 secret 5 $1$QxBT$i0o24EYUorGW8MGdm8.gE1
    username turnup-temp privilege 15 secret 5 $1$Mk4X$0a44f6jFsUke7lwozYCt5/
    username granitenoc secret 5 $1$J.jq$fKW6pxUQp.gCyuCxGz/lf0
    !
    redundancy
    bridge-domain 100 
    !
    !
    !         
    !
    !
    transceiver type all
    monitoring
    !
    ! 
    !
    crypto key generate rsa modulus 2048
    !
    !
    !
    !
    !
    ip domain-lookup
    ip name-server 8.8.8.8 4.2.2.2
    service call-home
    call-home
    contact-email-addr sch-smart-licensing@cisco.com
    profile "CiscoTAC-1"
    active
    destination transport-method http
    exit
    license smart enable
    !
    !
    ---------------------------------------------------------------------
    -----------------------*** BREAK COPY HERE ***-----------------------
    ---------------------------------------------------------------------
    !
    !
    interface GigabitEthernet0/0/0
    description  {cktId} {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface GigabitEthernet0/0/1
    description CUSTOMER LAN
    ip address {lanGateway} {lanSubnet}
    negotiation auto
    !
    !
    interface GigabitEthernet0/0/4
    description  {cktId} {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface GigabitEthernet0/0/5
    description CUSTOMER LAN
    ip address {lanGateway} {lanSubnet}
    negotiation auto
    !
    interface TenGigabitEthernet0/0/2
    description {cktId} // {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface TenGigabitEthernet0/0/3
    description OPTICAL CUSTOMER LAN
    no ip address
    !
    !
    interface TenGigabitEthernet0/0/12
    description {cktId} // {account}
    no ip address
    service instance 100 ethernet
    encapsulation untagged
    bridge-domain 100
    !
    !
    interface TenGigabitEthernet0/0/13
    description OPTICAL CUSTOMER LAN
    no ip address
    !
    !
    interface GigabitEthernet0
    vrf forwarding Mgmt-intf
    ip address 172.16.0.1 255.255.255.0
    negotiation auto
    !
    interface BDI100
    ip address  {wanIp} {wanSubnet}
    no shut
    !
    ip forward-protocol nd
    !
    ip bgp-community new-format
    no ip http server
    no ip http secure-server
    ip tftp source-interface GigabitEthernet0
    ip tacacs source-interface BDI100
    ip ssh time-out 60
    ip ssh authentication-retries 2
    ip ssh version 2
    ip ssh server algorithm encryption aes128-ctr aes192-ctr aes256-ctr aes128-cbc aes192-cbc aes256-cbc 3des-cbc
    ip ssh client algorithm encryption aes128-ctr aes192-ctr aes256-ctr aes128-cbc aes192-cbc aes256-cbc 3des-cbc
    ip route 0.0.0.0 0.0.0.0 {wanGateway}
    !
    ip access-list standard CPEAccess
    permit 162.223.83.42
    permit 162.223.83.38
    permit 65.202.145.2
    permit 134.204.10.2
    permit 134.204.11.2
    permit 172.85.135.238
    permit 172.85.180.240
    permit 72.46.171.2
    permit 172.16.0.0 0.15.255.255
    permit 10.0.0.0 0.255.255.255
    permit 192.168.0.0 0.0.255.255
    permit 198.18.0.0 0.1.255.255
    permit 198.51.100.0 0.0.0.255
    permit 100.64.1.0 0.0.0.255
    permit 100.64.2.0 0.0.0.255
    deny   any
    ip access-list standard GRANITESNMP
    permit 162.223.83.42
    permit 162.223.83.38
    permit 162.223.86.38
    permit 172.85.135.238
    permit 172.85.180.240
    permit 198.19.0.33
    permit 198.19.0.32
    permit 198.19.127.0 0.0.0.31
    deny   any log
    ip access-list standard GraniteNTP
    permit 162.223.83.36
    permit 162.223.86.36
    permit 198.18.0.0 0.1.255.255
    deny   any
    !
    !
    !
    snmp-server group SNMPv3Group v3 priv read SNMPv3View write SNMPv3View 
    snmp-server view SNMPv3View dod included
    SNMP-Server user SNMPv3User SNMPv3Group v3 auth md5 $$w1n$t@R!$$ priv des $$w1n$t@R!$$ access GRANITESNMP
    snmp-server community yNLQ14xxH4mgV RO GRANITESNMP
    snmp-server trap-source BDI100
    snmp-server location {location}
    snmp-server chassis-id {hostname}
    snmp-server enable traps bfd
    snmp-server enable traps config-copy
    snmp-server enable traps config
    snmp-server enable traps event-manager
    snmp-server enable traps cpu threshold
    snmp-server enable traps ethernet evc status create delete
    snmp-server enable traps alarms informational
    snmp-server enable traps ethernet cfm alarm
    snmp-server enable traps transceiver all

    snmp ifmib ifindex persist
    !
    tacacs-server directed-request
    !
    !
    !
    control-plane
    !
    banner motd ^CCCCC

    WARNING: To protect the system from unauthorized use and to ensure
    that the system is functioning properly, activities on this system are
    monitored and recorded and subject to audit.  Use of this system is
    expressed consent to such monitoring and recording.  Any unauthorized
    access or use of this Automated Information System is prohibited and
    could be subject to criminal and civil penalties.

    ^C
    !
    line con 0
    exec-timeout 15 0
    logging synchronous
    stopbits 1
    line vty 0 4
    access-class CPEAccess in vrf-also
    logging synchronous
    transport input ssh
    line vty 5 15
    access-class CPEAccess in vrf-also
    logging synchronous
    transport input ssh
    !         
    exception crashinfo file bootflash:crashinfo
    ntp source BDI100
    ntp access-group peer GraniteNTP
    ntp server 162.223.83.36 prefer
    ntp server 162.223.86.36
    !
    !
    end
                        """

                        elif configType == 'On-net_ASR-920':
                            write = f"""
    !
    no service pad
    service timestamps debug datetime localtime show-timezone
    service timestamps log datetime localtime show-timezone
    service password-encryption
    service sequence-numbers
    no platform punt-keepalive disable-kernel-core
    platform bfd-debug-trace 1
    platform xconnect load-balance-hash-algo mac-ip-instanceid
    platform tcam-parity-error enable
    platform tcam-threshold alarm-frequency 1
    !
    hostname {hostname}
    !
    vrf definition Mgmt-intf
    !
    address-family ipv4
    exit-address-family
    !
    address-family ipv6
    exit-address-family
    !
    logging buffered 51200 warnings
    enable secret 5 $1$bl2R$YvlGiFHmuA2R3FTeZhiZL/
    !
    aaa new-model
    !
    !
    aaa group server tacacs+ management
    server-private 198.19.0.49 timeout 1 key 7 013D0F400B5A5E5F70
    server-private 198.19.64.49 timeout 1 key 7 080F450A59485D4743
    ip tacacs source-interface Loopback66
    !
    aaa authentication login default group management local
    aaa authentication enable default group management enable
    aaa authorization console
    aaa authorization exec default group management local 
    aaa authorization commands 1 default group management local 
    aaa authorization commands 15 default group management local 
    aaa authorization network default group management local 
    aaa accounting exec default start-stop group management
    aaa accounting commands 1 default start-stop group management
    aaa accounting commands 15 default start-stop group management
    aaa accounting network default start-stop group management
    aaa accounting connection default start-stop group management
    aaa accounting system default start-stop group management
    !
    !
    !
    !
    !
    aaa session-id common
    clock timezone EST -5 0
    clock summer-time EDT recurring 2 Sun Mar 3:00 1 Sun Nov 3:00
    facility-alarm critical exceed-action shutdown
    !
    !
    !
    !         
    !
    !
    !
    !
    !



    no ip domain lookup
    ip domain name granitempls.com
    !
    ip dhcp pool Mgmt-intf
    network 172.16.0.0 255.255.255.0
    !
    !
    !
    login block-for 300 attempts 4 within 120
    login delay 2
    login on-failure log
    login on-success log
    !
    !
    !
    crypto key generate rsa mod 2048
    !
    !
    !
    !
    !
    !
    multilink bundle-name authenticated
    !
    !
    !
    sdm prefer default 
    !
    username NOCAdmin privilege 15 secret 5 $1$QxBT$i0o24EYUorGW8MGdm8.gE1
    username granitenoc secret 5 $1$J.jq$fKW6pxUQp.gCyuCxGz/lf0
    username turnup-temp privilege 15 secret 5 $1$Mk4X$0a44f6jFsUke7lwozYCt5/

    !
    redundancy
    bridge-domain 100 
    !
    !
    !
    !
    !
    transceiver type all
    monitoring
    !
    ! 
    !
    !
    !
    !
    ip domain-lookup
    ip name-server 8.8.8.8 4.2.2.2
    service call-home
    call-home
    contact-email-addr sch-smart-licensing@cisco.com
    profile "CiscoTAC-1"
    active
    destination transport-method http
    exit
    license smart enable
    !
    !
    !
    interface Loopback66
    description MGMT LB
    ip address {loopback66} 255.255.255.255
    !
    !
    interface GigabitEthernet0/0/0
    description {cktId} // {account}
    no ip address
    no shutdown
    negotiation auto
    service instance 100 ethernet
    {encapsulation}
    bridge-domain 100
    !
    !
    interface GigabitEthernet0/0/1
    description CUSTOMER LAN
    no ip address
    negotiation auto
    no shutdown
    service instance 200 ethernet
    encapsulation untagged
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    !
    interface GigabitEthernet0/0/4
    description {cktId} // {account}
    no ip address
    no shutdown
    negotiation auto
    service instance 100 ethernet
    {encapsulation}
    bridge-domain 100
    !
    !
    interface GigabitEthernet0/0/5
    description CUSTOMER LAN
    no ip address
    negotiation auto
    no shutdown
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    !
    interface TenGigabitEthernet0/0/12
    description {cktId} // {account}
    no ip address
    negotiation auto
    no shutdown
    service instance 100 ethernet
    {encapsulation}
    bridge-domain 100
    !
    !
    interface TenGigabitEthernet0/0/13
    description OPTICAL CUSTOMER LAN
    no ip address
    no shutdown
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    !
    !
    interface GigabitEthernet0
    vrf forwarding Mgmt-intf
    ip address 172.16.0.1 255.255.255.0
    negotiation auto
    !
    interface range GigabitEthernet0/0/2-23
    shut

    interface TenGigabitEthernet0/0/24
    description {cktId} // {account}
    no ip address
    service instance 100 ethernet
    {encapsulation}
    bridge-domain 100
    no shutdown


    interface TenGigabitEthernet0/0/25
    description // CUSTOMER_PUBLIC_LAN_Fiber//
    no ip address
    negotiation auto
    service instance 200 ethernet
    encapsulation untagged
    bridge-domain 200
    no shutdown

    interface range TenGigabitEthernet0/0/26-27
    shut
    !
    !
    !
    ---------------------------------------------------------------------
    -----------------------*** BREAK COPY HERE ***-----------------------
    ---------------------------------------------------------------------
    !
    !
    !
    interface BDI100
    ip address {wanIp} {wanSubnet}
    no shutdown
    !
    interface BDI200
    ip address {lanGateway} {lanSubnet}
    no shutdown
    !
    router bgp 65001
    bgp log-neighbor-changes
    neighbor {wanGateway} remote-as 16504
    !
    address-family ipv4
    network {loopback66} mask 255.255.255.255 route-map MGMT
    network {lanNetwork} mask 255.255.255.248
    neighbor {wanGateway} activate
    neighbor {wanGateway} send-community both
    exit-address-family
    !
    !
    !
    ip forward-protocol nd
    !
    ip bgp-community new-format
    ip community-list standard CO_METASWITCH_NO_IMPORT permit 16504:667
    ip community-list standard CO_EXTRANET_NO_ADVERTISE permit 16504:665
    no ip http server
    no ip http secure-server
    ip tftp source-interface GigabitEthernet0
    ip tacacs source-interface Loopback66
    ip ssh time-out 60
    ip ssh authentication-retries 2
    ip ssh version 2
    ip ssh server algorithm encryption aes128-ctr aes192-ctr aes256-ctr aes128-cbc aes192-cbc aes256-cbc 3des-cbc
    ip ssh client algorithm encryption aes128-ctr aes192-ctr aes256-ctr aes128-cbc aes192-cbc aes256-cbc 3des-cbc
    !
    !
    ip access-list standard CPEAccess
    permit 162.223.83.42
    permit 162.223.83.38
    permit 65.202.145.2
    permit 134.204.10.2
    permit 134.204.11.2
    permit 172.85.135.238
    permit 172.85.180.240
    permit 72.46.171.2
    permit 172.16.0.0 0.15.255.255
    permit 10.0.0.0 0.255.255.255
    permit 192.168.0.0 0.0.255.255
    permit 198.18.0.0 0.1.255.255
    permit 198.51.100.0 0.0.0.255
    permit 100.64.1.0 0.0.0.255
    permit 100.64.2.0 0.0.0.255
    permit 172.85.135.231
    permit 143.170.110.64 0.0.0.31
    permit 100.124.255.0 0.0.0.255
    deny   any
    ip access-list standard GRANITESNMP
    permit 162.223.83.42
    permit 162.223.83.38
    permit 162.223.86.38
    permit 172.85.135.238
    permit 172.85.180.240
    permit 198.19.0.33
    permit 198.19.0.32
    permit 198.19.127.0 0.0.0.31
    deny   any log
    ip access-list standard GraniteNTP
    permit 162.223.83.36
    permit 162.223.86.36
    permit 198.18.0.0 0.1.255.255
    deny   any
    ip access-list standard ZABBIXSNMP
    permit 198.19.127.64 0.0.0.31
    permit 198.19.127.0 0.0.0.31
    permit 198.19.127.96 0.0.0.31
    permit 198.19.127.128 0.0.0.31
    permit 198.19.128.0 0.0.1.255
    deny   any log
    !
    !
    !
    !
    ip prefix-list PFX_CE_LOOPBACKS seq 5 permit 198.18.0.0/15 ge 32
    !
    ip prefix-list PFX_VOICE_LOOPBACKS seq 10 permit 172.20.0.0/16 ge 32
    !
    route-map MGMT permit 10
    match ip address prefix-list PFX_CE_LOOPBACKS
    set community 16504:665 16504:12701
    !
    route-map MGMT permit 20
    match ip address prefix-list PFX_VOICE_LOOPBACKS
    set community 16504:665
    !
    route-map MGMT permit 30
    !
    route-map META_DROP permit 10
    set community 16504:667 additive
    !
    snmp-server group SNMPv3Group v3 priv read SNMPv3View write SNMPv3View 
    snmp-server view SNMPv3View dod included
    snmp-server community yNLQ14xxH4mgV RO GRANITESNMP
    snmp-server trap-source Loopback66
    snmp-server location {location}
    snmp-server chassis-id {hostname}
    snmp-server enable traps snmp authentication linkdown linkup coldstart warmstart
    snmp-server enable traps bgp
    snmp-server enable traps aaa_server
    snmp-server host 162.223.83.38 version 2c yNLQ14xxH4mgV 
    snmp-server host 162.223.83.42 version 2c yNLQ14xxH4mgV 
    snmp-server host 162.223.86.38 version 2c yNLQ14xxH4mgV 
    snmp-server host 198.19.0.32 version 2c yNLQ14xxH4mgV 
    snmp-server host 198.19.127.2 version 2c yNLQ14xxH4mgV
    snmp-server view ZABBIX mib-2 included
    snmp-server view ZABBIX sysUpTime.0 included
    snmp-server view ZABBIX ipTrafficStats included
    snmp-server view ZABBIX icmp included
    snmp-server view ZABBIX system.1.0 included
    snmp-server view ZABBIX ifIndex.1 included
    snmp-server view ZABBIX ifIndex.2 included
    snmp-server view ZABBIX ifIndex.3 included
    snmp-server view ZABBIX ifIndex.4 included
    snmp-server view ZABBIX ifIndex.5 included
    snmp-server view ZABBIX ifIndex.6 included
    snmp-server view ZABBIX ifIndex.* excluded
    snmp-server view ZABBIX snmpMIB.1.4.3.0 included
    snmp-server view ZABBIX snmpTraps.3 included
    snmp-server view ZABBIX snmpTraps.4 included
    snmp-server view ZABBIX 1.3.6.1.2.1* included
    snmp-server group zabbixgroup v3 priv read ZABBIX access ZABBIXSNMP
    snmp-server user zabbixuser zabbixgroup v3 auth sha lf3wzn9W8IuplwsoPn priv aes 128 lf3wzn9W8IuplwsoPn access ZABBIXSNMP
    !
    tacacs server PRIMARY
    address ipv4 198.19.0.49
    key 7 112710414743535C55
    timeout 1
    tacacs server SECONDARY
    address ipv4 198.19.64.49
    key 7 132B1E565B5D5C7A7A
    timeout 1
    !
    tacacs-server directed-request
    !
    !
    !
    control-plane
    !
    banner motd ^CCCC

    WARNING: To protect the system from unauthorized use and to ensure
    that the system is functioning properly, activities on this system are
    monitored and recorded and subject to audit.  Use of this system is
    expressed consent to such monitoring and recording.  Any unauthorized
    access or use of this Automated Information System is prohibited and
    could be subject to criminal and civil penalties.

    ^C
    !
    line con 0
    exec-timeout 15 0
    logging synchronous
    stopbits 1
    line vty 0 4
    access-class CPEAccess in vrf-also
    logging synchronous
    transport input ssh
    line vty 5 15
    access-class CPEAccess in vrf-also
    logging synchronous
    transport input ssh
    !
    exception crashinfo file bootflash:crashinfo

    ntp source Loopback66
    ntp access-group peer GraniteNTP
    ntp server 162.223.83.36 prefer
    ntp server 162.223.86.36
    !
    aaa authorization config-commands
    !
    !
    end
    """
                        else:
                            config_present = False
                            logging.warning(f"Config template not found for {tkt}. Please proceed manually.")
                        if config_present:
                            attach_file()


class VOIPRouterTemplate:
    def __init__(self):
        geolocator = Nominatim(user_agent='geoapiExercises')
        SMARTSHEET_ACCESS_TOKEN = os.getenv('SMARTSHEET_ACCESS_TOKEN')
        smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
        smart.errors_as_exceptions(True)
        sheet_id = 8329887249026948
        # Wrapped call to get the sheet
        sheet = smartsheet_api_call_with_retry(smart.Sheets.get_sheet, sheet_id)
        if sheet is None:
            raise Exception("Failed to fetch sheet data after retries")
        just_len = len("* Loaded Smartsheet: " + sheet.name)
        if just_len < 42:
            just_len = 42
        column_map = {column.title: column.id for column in sheet.columns}
        logging.warning("*" * (just_len + 2))
        logging.warning(("* Loaded Smartsheet: " + sheet.name).ljust(just_len) + " *")
        logging.warning(("* Compiling templates...".ljust(just_len) + " *"))
        logging.warning("*" * (just_len + 2))

        def get_value(x):
            zero_val = False
            value = row.get_column(column_map[x]).value
            if str(value).startswith("0"):
                zero_val = True
            try:
                value = float(value)
                if value.is_integer():
                    value = int(value)
            except:
                pass
            if zero_val:
                value = "0" + str(value)
            else:
                value = str(value)
            return value

        def update_cell(smart_column, smart_value):
            new_cell = smart.models.Cell()
            new_cell.column_id = column_map[smart_column]
            new_cell.value = smart_value
            new_cell.strict = False
            get_row = smart.models.Row()
            get_row.id = row.id
            get_row.cells.append(new_cell)
            response = smart.Sheets.update_rows(sheet_id, [get_row])

        class Attachments:
            def __init__(self, smartsheet_obj):
                """Init Attachments with base Smartsheet object."""
                self._base = smartsheet_obj
                self._log = logging.getLogger(__name__)

            def attach_file_to_row(self, sheet_id, row_id, _file):
                if not all(val is not None for val in ['sheet_id', 'row_id', '_file']):
                    raise ValueError(
                        ('One or more required values '
                         'are missing from call to ' + __name__))

                _op = fresh_operation('attach_file_to_row')
                _op['method'] = 'POST'
                _op['path'] = '/sheets/' + str(sheet_id) + '/rows/' + str(
                    row_id) + '/attachments'
                _op['files'] = {}
                _op['files']['file'] = _file
                _op['filename'] = _file

                expected = ['Result', 'Attachment']

                prepped_request = self._base.prepare_request(_op)
                response = self._base.request(prepped_request, expected, _op)

                return response

        def attach_file():
            newfile = f'{hostname}.txt'
            if json.loads(str(smart.Search.search_sheet(sheet_id=8892937224015748, query=newfile)))["totalCount"] == 0:
                with open(newfile, 'w') as f:
                    f.write(write)
                updated_attachment = smart.Attachments.attach_file_to_row(
                    sheet_id,  # sheet_id
                    row_id,  # row_id
                    (newfile,
                     open(newfile, 'rb'),
                     'application/notepad')
                )
                f.close()
                os.remove(newfile)
                logging.warning(f"{newfile} written!")
                update_cell('Template Status', 'Template attached')


            else:
                logging.warning(f"{newfile} already exists!")

        def extract_state(location):
            parts = location.rsplit(',', 1)
            return parts[1].strip() if len(parts) > 1 else None

        for row in sheet.rows:
            row_id = row.id
            if True:
                tkt = normalize_ticket_number(get_value('Equipment Ticket'))
                status = get_value('Status')
                if status in ('Allocated', 'In Process', 'None', 'Pending Information'):
                    account = get_value('Child Account')
                    hostname = get_value('Hostname')
                    cktId = get_value('Circuit ID')
                    wanIp = get_value('WAN Usable IP')
                    wanSubnet = get_value('WAN Subnet Mask')
                    wanGateway = get_value('WAN Gateway IP')
                    city_state = str(get_value("City, State/Province"))
                    city_state = usaddress.tag(city_state)
                    city_state = json.loads(json.dumps(city_state[0]))
                    try:
                        state = city_state['StateName']
                    except:
                        state = extract_state(str(get_value("City, State/Province")))
                    if state in (
                            'ME', 'NH', 'VT', 'MA', 'RI', 'CT', 'NY', 'NJ', 'PA', 'DE', 'MD', 'WV', 'DC', 'VA', 'NC',
                            'SC', 'GA', 'FL', 'AL', 'TN', 'KY', 'OH', 'MI', 'IN', 'MI', 'IL', 'WI', 'MS'):
                        # East
                        nameserver = '162.223.83.36 162.223.86.36'
                        sip_domain = 'aosnyc.granitevoip.com'
                    else:
                        # West
                        nameserver = '162.223.86.36 162.223.83.36'
                        sip_domain = 'aoslax.granitevoip.com'
                    number_of_paths = int(get_value('Number of Paths'))
                    if number_of_paths == 23:
                        pri_paths = '1-24'
                    else:
                        pri_paths = f'1-{number_of_paths},24'
                    sip_username = get_value('SIP Username')
                    sip_password = get_value('SIP Password')

                    # Check for missing variables in smartsheet
                    varCheck = {"Equipment Ticket": tkt,
                                "Child Account": account,
                                "Hostname": hostname,
                                "Circuit ID": cktId,
                                "WAN Subnet Mask": wanSubnet,
                                "WAN Gateway IP": wanGateway,
                                "WAN Usable IP": wanIp,
                                "Number of Paths": number_of_paths,
                                "City, State/Province": city_state}
                    varList = []
                    missingData = False
                    for var in varCheck:
                        if varCheck[var] in ["None", "none", "n/a", "na", "N/A", "NA"] or varCheck[var] is None:
                            missingData = True
                            varList.append(var)
                    if missingData:
                        logging.warning(
                            textwrap.fill(f'{tkt} missing data in column(s): {varList}', width=(just_len + 2)))
                        var_notes = str(varList).replace("[", "").replace("]", "").replace("'", "")
                        missing_notes = f'\nMissing data: {var_notes}'
                        update_cell('Notes', missing_notes)
                        update_cell('Template Status', 'Missing data')
                        update_cell('Status', 'Pending Information')
                        config_present = False
                    else:
                        config_present = True
                        write = f"""
!
!
hostname {hostname}
enable password Granite1!
!
!
name-server {nameserver}
!
!
!
ip subnet-zero
ip classless
ip routing
no ipv6 unicast-routing
!
!
!
!
no auto-config
auto-config authname adtran encrypted password 2d2b81a644467cd85106be31cf767d5bc500
!
event-history on
no logging forwarding
no logging email
!
service password-encryption
!
username "aweiner" password encrypted "3b3c7280764f1c3a2cd961202185124bef3f" 
username "granitenoc" password encrypted "1418a9fc9d359e4fae862d2f923c5aef8ed3" 
username "smcelroy" password encrypted "2621c35600e562d4ea15d5cace18e9162d35" 
username "turnup-temp" password encrypted "2926704f29cfc5d16e3152bc5885dfe3eacd" 
username "turnup-temp" privilege 7 password encrypted "2926704f29cfc5d16e3152bc5885dfe3eacd" 
username "turnup-temp" privilege 15 password encrypted "2926704f29cfc5d16e3152bc5885dfe3eacd"
username "NOCAdmin" privilege 7 password encrypted "242ee79a6e28c8aa9bc6834b14ab5e607014"
!
!
!
aaa on
aaa processes 3
!
tacacs-server timeout 1
tacacs-server host 172.85.135.235 timeout 1 key V#38fe;5K[

!
aaa authentication login local local
aaa authentication login vty group tacacs+ local
!
aaa authentication enable default enable
!
aaa authorization console
aaa authorization commands 1 local if-authenticated
aaa authorization commands 1 vty group tacacs+ if-authenticated
aaa authorization commands 15 local if-authenticated
aaa authorization commands 15 vty group tacacs+ if-authenticated
aaa authorization exec local if-authenticated
aaa authorization exec vty group tacacs+ if-authenticated
aaa accounting suppress null-username
aaa accounting commands 1 vty stop-only group tacacs+
aaa accounting commands 1 local stop-only group tacacs+
aaa accounting commands 15 vty stop-only group tacacs+
aaa accounting commands 15 local stop-only group tacacs+
aaa accounting connection vty start-stop group tacacs+
aaa accounting connection local start-stop group tacacs+
aaa accounting exec vty start-stop group tacacs+
aaa accounting exec local start-stop group tacacs+
!
!
!
banner motd ^
WARNING: To protect the system from unauthorized use and to ensure
that the system is functioning properly, activities on this system are
monitored and recorded and subject to audit.  Use of this system is
expressed consent to such monitoring and recording.  Any unauthorized
access or use of this Automated Information System is prohibited and
could be subject to criminal and civil penalties.
^
!
!
ip firewall
no ip firewall alg msn
no ip firewall alg mszone
no ip firewall alg h323
!
!
!
!
!
!
!
!
no dot11ap access-point-control
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
qos map VOIP 10
  match dscp ef
  priority percent 80
qos map VOIP 20
  match dscp af31 cs3
  set dscp cs4
qos map VOIP 30
  match any
  set dscp default
!
!
!
!
interface eth 0/1
  description [ {cktId} ] [ {account} ]
  ip address  {wanIp} {wanSubnet}
  media-gateway ip primary 
  no shutdown
!
!
interface eth 0/2
  no ip address
  shutdown
!
!
!
interface gigabit-eth 0/1
  no ip address
  shutdown
!
!
!
!
interface t1 0/1
  description // not in use //
  shutdown
!
interface t1 0/2
  description // not in use //
  shutdown
!
interface t1 0/3
  description // Physical PRI to Customer PBX //
  tdm-group 1 timeslots {pri_paths} speed 64
  no shutdown
!
interface t1 0/4
  description // not in use //
  shutdown
!
!
interface pri 1
  description // Logical PRI interface 1 to Customer PBX //
  isdn name-delivery display
  connect t1 0/3 tdm-group 1
  no shutdown
!
!
interface fxs 0/1
  shutdown
!
interface fxs 0/2
  shutdown
!
interface fxs 0/3
  shutdown
!
interface fxs 0/4
  shutdown
!
interface fxs 0/5
  shutdown
!
interface fxs 0/6
  shutdown
!
interface fxs 0/7
  shutdown
!
interface fxs 0/8
  shutdown
!
!
!
!
isdn-group 1
  connect pri 1
!
!
!
!
!
!
!
!
!
!
!
ip route 0.0.0.0 0.0.0.0 {wanGateway}
!
no tftp server
no tftp server overwrite
no http server
no http secure-server
no ip ftp server
no ip scp server
no ip sntp server
!
!
!
!
ip access-list standard GRANITESIPSERVERS
  permit 8.38.33.192 0.0.0.63
  permit 8.42.0.192 0.0.0.63
  permit 8.42.1.192 0.0.0.63
  permit 63.210.209.192 0.0.0.63
  permit 162.223.83.224 0.0.0.31
  permit 162.223.86.224 0.0.0.31
  permit 162.223.87.240 0.0.0.15
!
ip access-list standard GRANITESNMP
  permit host 198.19.0.32
  permit host 198.19.0.33
  permit host 172.85.135.238
  permit host 172.85.180.240
  permit host 198.19.0.32
  permit host 198.19.0.33
  permit host 162.223.83.38
  permit host 162.223.86.38
  permit host 162.223.83.42
  permit 198.19.127.0 0.0.0.31
  deny   any log
!
ip access-list standard CPEAccess
  permit 172.16.0.0 0.15.255.255
  permit 10.0.0.0 0.255.255.255
  permit 192.168.0.0 0.0.255.255
  permit 198.18.0.0 0.1.255.255
  permit 198.51.100.0 0.0.0.255
  permit 100.64.1.0 0.0.0.255
  permit 100.64.2.0 0.0.0.255
  permit 198.19.127.0 0.0.0.31
  permit 100.124.255.192 0.0.0.63
  permit host 65.202.145.2
  permit host 172.85.228.254
  permit host 134.201.10.2
  permit host 134.201.11.2
  permit host 72.46.171.2
  permit host 172.85.135.238
  permit host 172.85.180.240
  permit host 162.223.83.38
  permit host 162.223.83.42
  deny   any
!
ip access-list standard NTPAccess
  permit host 162.223.83.36
  permit host 162.223.86.36
  deny   any
!					   
!
!
snmp agent
snmp-server group SNMPv3Group v3 priv read SNMPv3View write SNMPv3View 
snmp-server view SNMPv3View 1.3.6 included 
SNMP-Server user SNMPv3User SNMPv3Group v3 auth md5 $$w1n$t@R!$$ priv des $$w1n$t@R!$$
snmp-server host 172.85.135.238 informs version 3 auth SNMPv3Group
snmp-server host 172.85.180.240 informs version 3 auth SNMPv3Group
!
auto-link
auto-link server primary 172.85.135.230
!
!
ip rtp quality-monitoring
ip rtp quality-monitoring sip
ip rtp quality-monitoring history max-streams 50
!
ip rtp quality-monitoring reporter "ReporterVQM"
  description "VQM To n-Commmand"
  collector auto-link
  no shutdown
!
sip
sip udp 5060
no sip tcp
no sip tls
!
!
!
voice feature-mode network
voice flashhook mode transparent
voice forward-mode network
!
!
!
!
!
!
!
!
voice dial-plan 1 long-distance 1-NXX-NXX-XXXX
voice dial-plan 2 long-distance NXX-NXX-XXXX
!
!
!
!
voice codec-list TRUNK
  codec g711ulaw
  codec g729
!
!
!
voice trunk T01 type isdn
  description "PRI"
  caller-id-override number-inbound {sip_username} if-no-cpn
  resource-selection linear descending
  connect isdn-group 1
  no early-cut-through
  rtp delay-mode adaptive
  rtp qos dscp 46
!
!
voice trunk T02 type sip
  sip-server primary {sip_domain}
  registrar primary {sip_domain}
  registrar threshold absolute 15
  outbound-proxy primary {sip_domain}
  domain "{sip_domain}"
  dial-string source to
  register {sip_username} auth-name {sip_username} password  {sip_password}
  codec-list TRUNK both
  authentication username {sip_username} password {sip_password}
  match ani "NXXX" add p-asserted-identity {sip_username}
  match ani "NXXXXXX" add p-asserted-identity {sip_username}
  match ani "NXXXXXXXXX" add p-asserted-identity {sip_username}
  match ani "1NXXXXXXXXX" add p-asserted-identity {sip_username}
!
!
voice grouped-trunk SIP_GROUPED_TRUNK
  trunk T02
  accept NXX-XXXX cost 0
  accept 1-NXX-XXX-XXXX cost 0
  accept $ cost 0
  accept NXX-XXX-XXXX cost 0
!
!
voice grouped-trunk ISDN_GROUPED_TRUNK
  trunk T01
  accept NXX-XXXX cost 0
  accept 1-NXX-XXX-XXXX cost 0
  accept $ cost 0
  accept NXX-XXX-XXXX cost 0
  accept XXXX cost 0
!
!
!
!
!
!
sip access-class ip "GRANITESIPSERVERS" in
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
!
line con 0
  login authentication vty
  authorization commands 1 vty
  authorization commands 15 vty
  authorization commands 7 vty
  authorization exec vty
!
line telnet 0 4
  shutdown
!
line ssh 0 4
  login authentication vty
  authorization commands 1 vty
  authorization commands 15 vty
  authorization exec vty
  no shutdown
  authorization commands 7 vty
  ip access-class CPEAccess in
!
!
ntp source ethernet 0/1
ntp ip access-class NTPAccess in
ntp server ntp.granitevoip.com
!
!
!
end
"""
                    if config_present:
                        attach_file()
                    logging.warning("-" * (just_len + 2))
        logging.warning("*" * (just_len + 2))


class MergedVoice:
    def __init__(self):
        sheet_id = 8892937224015748

        SMARTSHEET_ACCESS_TOKEN = os.getenv('SMARTSHEET_ACCESS_TOKEN')
        smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
        smart.errors_as_exceptions(True)
        # Wrapped call to get the sheet
        sheet = smartsheet_api_call_with_retry(smart.Sheets.get_sheet, sheet_id)
        if sheet is None:
            raise Exception("Failed to fetch sheet data after retries")
        just_len = len("* Loaded Smartsheet: " + sheet.name)
        if just_len < 42:
            just_len = 42
        logging.warning("*" * (just_len + 2))
        column_map = {column.title: column.id for column in sheet.columns}

        try:
            connection = pymssql.connect(server='gp2018', user=f'GRT0\\{os.getenv('GRT_USER')}',
                                         password=os.getenv('GRT_PASS'),
                                         database='SBM01')
        except Error as e:
            logging.warning(str(e).replace("\\n", "\n"))
            sys.exit()
        logging.warning('* Connected to SQL Server: GP2018.SBM01'.ljust(just_len) + " *")
        logging.warning("* Adding allocated Voice tickets...".ljust(just_len) + " *")
        logging.warning("*" * (just_len + 2))
        time.sleep(1)
        cursor = connection.cursor()

        def get_value(x):
            zero_val = False
            value = row.get_column(column_map[x]).value
            if str(value).startswith("0"):
                zero_val = True
            try:
                value = float(value)
                if value.is_integer():
                    value = int(value)
            except:
                pass
            if zero_val:
                value = "0" + str(value)
            else:
                value = str(value)
            return value

        def add_row(ticket_num, account_num, customer_name):
            new_row = smart.models.Row()
            new_row.to_bottom = True
            new_row.cells.append({
                'column_id': column_map['Equipment Ticket'],
                'value': ticket_num
            })
            new_row.cells.append({
                'column_id': column_map['Child Account'],
                'value': account_num
            })
            new_row.cells.append({
                'column_id': column_map['Status'],
                'value': 'Allocated'
            })
            new_row.cells.append({
                'column_id': column_map['Customer Name'],
                'value': customer_name
            })
            new_row.cells.append({
                'column_id': column_map['Requested Ship'],
                'value': req_ship
            })
            new_row.cells.append({
                'column_id': column_map['Equipment Type'],
                'value': 'Algo/ATA/Phones'
            })
            # Wrap the add_rows call in a lambda to defer execution
            add_rows_call = lambda: smart.Sheets.add_rows(sheet_id, [new_row])

            # Use the retry function to attempt to add the row
            response = smartsheet_api_call_with_retry(add_rows_call)

        def update_cell(smart_column, smart_value):
            new_cell = smart.models.Cell()
            new_cell.column_id = column_map[smart_column]
            new_cell.value = smart_value
            new_cell.strict = False
            get_row = smart.models.Row()
            get_row.id = row.id
            get_row.cells.append(new_cell)
            # Wrap the update_rows call in a lambda to defer execution
            update_rows_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

            # Use the retry function to attempt to update the row
            response = smartsheet_api_call_with_retry(update_rows_call)

        def pull_data():
            cursor.execute(f"""
            SELECT DISTINCT SOP10200.SOPNUMBE as Ticket, SOP10100.CSTPONBR as Account, SOP10100.CUSTNAME as Customer, SOP10100.ReqShipDate
            FROM SOP10200
            FULL JOIN SOP10100 ON SOP10200.SOPNUMBE = SOP10100.SOPNUMBE
            FULL JOIN SOP10106 ON SOP10200.SOPNUMBE = SOP10106.SOPNUMBE
            FULL JOIN IV00101 ON SOP10200.ITEMNMBR = IV00101.ITEMNMBR
            FULL JOIN IV00102 ON SOP10200.ITEMNMBR = IV00102.ITEMNMBR
            FULL JOIN SOP10201 ON SOP10200.ITEMNMBR = SOP10201.ITEMNMBR AND SOP10200.SOPNUMBE = SOP10201.SOPNUMBE
            WHERE
            SOP10200.SOPTYPE = '2'
            AND SOP10201.SERLTNUM NOT LIKE '%*%'
        	AND SOP10200.ITEMNMBR IN ({devices})    
        	AND SOP10100.BACHNUMB = 'VOIP CONFIG'     
        ORDER BY SOP10200.SOPNUMBE



            """)

        pull_data()
        sql_row = cursor.fetchone()
        sqlRowCount = 0
        sql_list = []
        while sql_row:
            sqlRowCount += 1
            sql_list.append(sqlRowCount)
            sql_row = cursor.fetchone()
        pull_data()
        for i in sql_list:
            sql_row = cursor.fetchone()
            tkt = sql_row[0].strip()
            acct = sql_row[1].strip()
            cust_name = sql_row[2].strip()
            req_ship = datetime.datetime.strptime(str(sql_row[3]), "%Y-%m-%d %H:%M:%S").isoformat()[:-3] + 'Z'
            tkt_exists = False
            for row in sheet.rows:
                if tkt.replace("CW", "").replace("-1", "") == get_value('Equipment Ticket').replace("CW", "").replace(
                        "-1", ""):
                    tkt_exists = True
            if not tkt_exists:
                add_row(tkt, acct, cust_name)
                logging.warning(f'Ticket: [{tkt}, {acct}, {cust_name}] added!')
                logging.warning("-" * (just_len + 2))
                time.sleep(2)


class Edgeboot:
    def __init__(self):
        sheet_id = 8892937224015748

        SMARTSHEET_ACCESS_TOKEN = os.getenv('SMARTSHEET_ACCESS_TOKEN')
        smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
        smart.errors_as_exceptions(True)
        # Wrapped call to get the sheet
        sheet = smartsheet_api_call_with_retry(smart.Sheets.get_sheet, sheet_id)
        if sheet is None:
            raise Exception("Failed to fetch sheet data after retries")
        just_len = len("* Loaded Smartsheet: " + sheet.name)
        if just_len < 42:
            just_len = 42
        logging.warning("*" * (just_len + 2))
        column_map = {column.title: column.id for column in sheet.columns}

        try:
            connection = pymssql.connect(server='gp2018', user=f'GRT0\\{os.getenv('GRT_USER')}',
                                         password=os.getenv('GRT_PASS'),
                                         database='SBM01')
        except Error as e:
            logging.warning(str(e).replace("\\n", "\n"))
            sys.exit()
        logging.warning('* Connected to SQL Server: GP2018.SBM01'.ljust(just_len) + " *")
        logging.warning("* Adding allocated Edgeboot tickets...".ljust(just_len) + " *")
        logging.warning("*" * (just_len + 2))
        time.sleep(1)
        cursor = connection.cursor()

        def get_value(x):
            zero_val = False
            value = row.get_column(column_map[x]).value
            if str(value).startswith("0"):
                zero_val = True
            try:
                value = float(value)
                if value.is_integer():
                    value = int(value)
            except:
                pass
            if zero_val:
                value = "0" + str(value)
            else:
                value = str(value)
            return value

        def add_row(ticket_num, account_num, customer_name, req_ship, location, notes):
            new_row = smart.models.Row()
            new_row.to_bottom = True
            new_row.cells.append({
                'column_id': column_map['Equipment Ticket'],
                'value': ticket_num
            })
            new_row.cells.append({
                'column_id': column_map['Child Account'],
                'value': account_num
            })
            new_row.cells.append({
                'column_id': column_map['Status'],
                'value': 'Allocated'
            })
            new_row.cells.append({
                'column_id': column_map['Customer Name'],
                'value': customer_name
            })
            new_row.cells.append({
                'column_id': column_map['Requested Ship'],
                'value': req_ship
            })
            new_row.cells.append({
                'column_id': column_map['Equipment Type'],
                'value': 'Edgeboot'
            })
            new_row.cells.append({
                'column_id': column_map['Prov Username'],
                'value': 'ayetman'
            })
            new_row.cells.append({
                'column_id': column_map['City, State/Province'],
                'value': location
            })
            new_row.cells.append({
                'column_id': column_map['Notes'],
                'value': notes
            })

            # Wrap the add_rows call in a lambda to defer execution
            add_rows_call = lambda: smart.Sheets.add_rows(sheet_id, [new_row])

            # Use the retry function to attempt to add the row
            response = smartsheet_api_call_with_retry(add_rows_call)

        def update_cell(smart_column, smart_value):
            new_cell = smart.models.Cell()
            new_cell.column_id = column_map[smart_column]
            new_cell.value = smart_value
            new_cell.strict = False
            get_row = smart.models.Row()
            get_row.id = row.id
            get_row.cells.append(new_cell)
            # Wrap the update_rows call in a lambda to defer execution
            update_rows_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

            # Use the retry function to attempt to update the row
            response = smartsheet_api_call_with_retry(update_rows_call)

        def pull_data():
            cursor.execute(f"""
SELECT DISTINCT SOP10200.SOPNUMBE as Ticket, SOP10100.CSTPONBR as Account, SOP10100.CUSTNAME as Customer, SOP10100.ReqShipDate, SOP10100.CITY as City, SOP10100.STATE as 'State'

FROM SOP10200
FULL JOIN SOP10100 ON SOP10200.SOPNUMBE = SOP10100.SOPNUMBE
FULL JOIN SOP10106 ON SOP10200.SOPNUMBE = SOP10106.SOPNUMBE
FULL JOIN IV00101 ON SOP10200.ITEMNMBR = IV00101.ITEMNMBR
FULL JOIN IV00102 ON SOP10200.ITEMNMBR = IV00102.ITEMNMBR
FULL JOIN SOP10201 ON SOP10200.ITEMNMBR = SOP10201.ITEMNMBR AND SOP10200.SOPNUMBE = SOP10201.SOPNUMBE
WHERE
SOP10200.SOPTYPE = '2'
AND SOP10201.SERLTNUM NOT LIKE '%*%'
AND SOP10200.ITEMNMBR IN ('EDGEBOOT', 'EDGEBOOT+')    
AND SOP10100.BACHNUMB IN ('MOBILITY CONFIG', 'CONFIG LAB', 'VOIP CONFIG') 
ORDER BY SOP10200.SOPNUMBE
            """)

        pull_data()
        sql_row = cursor.fetchone()
        sqlRowCount = 0
        sql_list = []
        while sql_row:
            sqlRowCount += 1
            sql_list.append(sqlRowCount)
            sql_row = cursor.fetchone()
        pull_data()
        for i in sql_list:
            sql_row = cursor.fetchone()
            tkt = sql_row[0].strip()
            acct = sql_row[1].strip()
            cust_name = sql_row[2].strip()
            req_ship = datetime.datetime.strptime(str(sql_row[3]), "%Y-%m-%d %H:%M:%S").isoformat()[:-3] + 'Z'
            location = f"{sql_row[4].strip()}, {sql_row[5].strip()}"
            tkt_exists = False
            for row in sheet.rows:
                if tkt.replace("CW", "").replace("-1", "") == get_value('Equipment Ticket').replace("CW", "").replace(
                        "-1", ""):
                    tkt_exists = True
            if not tkt_exists:
                try:
                    notes = "***" + str(json.loads(str(GetCWInfo(normalize_ticket_number(tkt))))["Sub-Type"]) + "***"
                except Exception:
                    notes = "Standard Edgeboot (Non-MNS)"
                add_row(tkt, acct, cust_name, req_ship, location, notes)
                logging.warning(f'Ticket: [{tkt}, {acct}, {cust_name}, {location}, {notes}] added!')
                logging.warning("-" * (just_len + 2))
                time.sleep(2)


class ClearHelpers:
    def __init__(self):
        SMARTSHEET_ACCESS_TOKEN = os.getenv('SMARTSHEET_ACCESS_TOKEN')
        smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
        smart.errors_as_exceptions(True)
        sheet_ids = [

            4813429487390596,
            1556213305134980,
            1153412976562052,
            7785735834783620,
            8031837460844420,
            468271203569540,
            2487927766470532,
            7810199431630724,
            3517946554961796,
            8208669854355332,
            7165863471828868,
            1750520939106180

        ]
        for sheet_id in sheet_ids:
            sheet = smartsheet_api_call_with_retry(smart.Sheets.get_sheet, sheet_id)
            if sheet is None:
                raise Exception("Failed to fetch sheet data after retries")
            column_map = {column.title: column.id for column in sheet.columns}
            sheet_name = sheet.name
            just_len = max(len("* Clearing " + sheet_name), 42)
            logging.warning("*" * (just_len + 2))
            logging.warning(("* Clearing " + sheet_name).ljust(just_len) + " *")
            rows_to_delete = [row.id for row in sheet.rows]
            for x in range(0, len(rows_to_delete), 100):
                result = smartsheet_api_call_with_retry(smart.Sheets.delete_rows, sheet_id, rows_to_delete[x:x + 100])
                if result is None:
                    logging.error(f"Failed to delete rows in sheet {sheet_name} after retries")


class TrackingUpdate:
    def __init__(self, sheet_id, ticket_column, tracking_column=None, tracking_status_column=None, del_date_column=None,
                 tkt=None):
        self.sheet_id = sheet_id
        self.ticket_column = ticket_column
        self.tracking_column = tracking_column
        self.status_column = tracking_status_column

        def get_value(x):
            zero_val = False
            value = row.get_column(column_map[x]).value
            if str(value).startswith("0"):
                zero_val = True
            try:
                value = float(value)
                if value.is_integer():
                    value = int(value)
            except:
                pass
            if zero_val:
                value = "0" + str(value)
            else:
                value = str(value)
            return value

        def recognize_delivery_service(tracking):
            service = None

            usps_pattern = [
                '^(94|93|92|94|95)[0-9]{20}$',
                '^(94|93|92|94|95)[0-9]{22}$',
                '^(70|14|23|03)[0-9]{14}$',
                '^(M0|82)[0-9]{8}$',
                '^([A-Z]{2})[0-9]{9}([A-Z]{2})$'
            ]

            ups_pattern = [
                '^(1Z)[0-9A-Z]{16}$',
                '^(T)+[0-9A-Z]{10}$',
                '^[0-9]{8}$',
                '^[0-9]{9}$',
                '^[0-9]{26}$'
            ]

            fedex_pattern = [
                '^[0-9]{20}$',
                '^[0-9]{15}$',
                '^[0-9]{12}$',
                '^[0-9]{22}$'
            ]

            usps = "(" + ")|(".join(usps_pattern) + ")"
            fedex = "(" + ")|(".join(fedex_pattern) + ")"
            ups = "(" + ")|(".join(ups_pattern) + ")"

            if re.match(usps, tracking) is not None:
                service = 'USPS'
            elif re.match(ups, tracking) is not None:
                service = 'UPS'
            elif re.match(fedex, tracking) is not None:
                service = 'FedEx'

            return service

        SMARTSHEET_ACCESS_TOKEN = os.getenv('SMARTSHEET_ACCESS_TOKEN')
        smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
        smart.errors_as_exceptions(True)
        # Wrapped call to get the sheet
        sheet = smartsheet_api_call_with_retry(smart.Sheets.get_sheet, sheet_id)
        if sheet is None:
            raise Exception("Failed to fetch sheet data after retries")
        just_len = len("* Loaded Smartsheet: " + sheet.name)
        if just_len < 42:
            just_len = 42
        logging.warning("*" * (just_len + 2))
        logging.warning(("* Loaded Smartsheet: " + sheet.name).ljust(just_len) + " *")
        column_map = {column.title: column.id for column in sheet.columns}
        try:
            connection = pymssql.connect(server='gp2018', user=f'GRT0\\{os.getenv('GRT_USER')}',
                                         password=os.getenv('GRT_PASS'),
                                         database='SBM01')
            logging.warning('* Connected to SQL Server: GP2018.SBM01'.ljust(just_len) + " *")
            cursor = connection.cursor()
        except Error as e:
            logging.warning(str(e).replace("\\n", "\n"))
            exit()
        logging.warning("*" * (just_len + 2))
        logging.warning("-" * (just_len + 2))
        session = requests.session()
        ups_headers = {
            'accept': 'application/json',
            'Authorization': os.getenv('UPS_ACCESS_TOKEN'),
            'Content-Type': 'application/x-www-form-urlencoded',
        }

        ups_data = {
            'grant_type': 'client_credentials',
        }

        ups_access_token = json.loads(session.post(
            'https://wwwcie.ups.com/security/v1/oauth/token', headers=ups_headers, data=ups_data).content)[
            "access_token"]

        # get fedex access_token
        fedex_payload = {
            'grant_type': "client_credentials",
            'client_id': os.getenv('FEDEX_CLIENT_ID'),
            'client_secret': os.getenv('FEDEX_CLIENT_SECRET'),
        }
        fedex_headers = {
            'Content-Type': "application/x-www-form-urlencoded",
        }

        fedex_access_token = json.loads(session.post("https://apis.fedex.com/oauth/token", data=fedex_payload,
                                                     headers=fedex_headers).text)['access_token']

        for row in sheet.rows:
            ss_date = False
            if get_value(tracking_column) is False or get_value(
                    tracking_status_column).lower().strip() not in (
                    'ticket not found', 'delivered', 'returned to sender',
                    'rdy to invoice - no tracking available. this may be a partial shipment. tracking may be '
                    'associated with another ticket number.', 'we could not locate the shipment details for this '
                                                              'tracking number. details are only available for '
                                                              'shipments made within the last 120 days. please check '
                                                              'your information.'):
                if get_value(ticket_column):
                    tkt = normalize_ticket_number(get_value(ticket_column).strip())
                    cursor.execute(f"""DECLARE @TicketNumber NVARCHAR(100) = '{tkt}';

SELECT DISTINCT *
FROM (
    SELECT DISTINCT
        sop30300.SOPNUMBE AS 'Equipment Ticket', 
        sop10107.Tracking_Number,
        CASE
            WHEN sop30200.BACHNUMB IN ('RDY TO INVOICE', 'RDY TO INV', 'Q', 'Q1', 'Q2', 'Q3', 'Q4') THEN 'RDY TO INVOICE'
            ELSE sop30200.BACHNUMB
        END AS 'Queue'
    FROM sop30300
    LEFT JOIN sop30200 ON sop30200.SOPNUMBE = sop30300.SOPNUMBE
    LEFT JOIN sop10107 ON sop10107.SOPNUMBE = sop30300.SOPNUMBE
    WHERE sop30300.SOPNUMBE = @TicketNumber
       OR sop30300.SOPNUMBE = @TicketNumber + '.1'
       OR sop30300.SOPNUMBE = @TicketNumber + '.2'
       OR sop30300.SOPNUMBE = @TicketNumber + '.3'
       OR sop30300.SOPNUMBE = 'CW' + @TicketNumber
       OR sop30300.SOPNUMBE = 'CW' + @TicketNumber + '-1'
       OR sop30300.SOPNUMBE = 'CW' + @TicketNumber + '-2'
       OR sop30300.SOPNUMBE = 'CW' + @TicketNumber + '-3'

    UNION ALL

    SELECT DISTINCT
        sop10100.SOPNUMBE AS 'Equipment Ticket', 
        sop10107.Tracking_Number,
        CASE
            WHEN sop10100.BACHNUMB IN ('RDY TO INVOICE', 'RDY TO INV', 'Q', 'Q1', 'Q2', 'Q3', 'Q4') THEN 'RDY TO INVOICE'
            ELSE sop10100.BACHNUMB
        END AS 'Queue'
    FROM sop10100
    LEFT JOIN sop10107 ON sop10107.SOPNUMBE = sop10100.SOPNUMBE
    WHERE sop10100.SOPNUMBE = @TicketNumber
       OR sop10100.SOPNUMBE = @TicketNumber + '.1'
       OR sop10100.SOPNUMBE = @TicketNumber + '.2'
       OR sop10100.SOPNUMBE = @TicketNumber + '.3'
       OR sop10100.SOPNUMBE = 'CW' + @TicketNumber
       OR sop10100.SOPNUMBE = 'CW' + @TicketNumber + '-1'
       OR sop10100.SOPNUMBE = 'CW' + @TicketNumber + '-2'
       OR sop10100.SOPNUMBE = 'CW' + @TicketNumber + '-3'
) AS MASTER_GP_QUERY;
""")
                    try:
                        sql_row = cursor.fetchone()
                        inquiry_number = sql_row[1].strip().upper()
                    except:
                        inquiry_number = False
                    if inquiry_number:
                        shipper = recognize_delivery_service(inquiry_number)
                        if shipper == 'UPS':
                            headers = {
                                'accept': '*/*',
                                'transId': tkt,
                                'transactionSrc': 'testing',
                                'Authorization': f'Bearer {ups_access_token}',
                                'Content-Type': 'application/json',
                            }

                            params = {
                                'locale': 'en_US',
                                'returnSignature': 'false',
                            }

                            tracking_info = session.get(
                                f'https://onlinetools.ups.com/api/track/v1/details/{inquiry_number}', params=params,
                                headers=headers).content
                            try:
                                status = json.loads(
                                    tracking_info)['trackResponse']['shipment'][0]['package'][0]['activity'][0][
                                    'status'][
                                    'description'].strip()
                                try:
                                    real_date = False
                                    ss_date = False
                                    del_date = \
                                        json.loads(tracking_info)['trackResponse']['shipment'][0]['package'][0][
                                            'deliveryDate'][
                                            0][
                                            'date']
                                    del_date = datetime.datetime.strptime(del_date, '%Y%m%d')
                                    real_date = del_date.isoformat()[:-3] + 'Z'
                                    # logging.warning(real_date)
                                    ss_date = datetime.datetime.strftime(del_date, "%Y-%m-%d")
                                    del_date = f"{str(calendar.month_name[del_date.month])} {str(del_date.day)}," \
                                               f" {str(del_date.year)}".strip()
                                    del_time = \
                                        json.loads(tracking_info)['trackResponse']['shipment'][0]['package'][0][
                                            'deliveryTime'][
                                            'type'].strip()
                                    if del_time == "CMT":
                                        del_time = "AM"
                                    datestr = f"{del_date} by {del_time}"
                                except:
                                    datestr = "The delivery date will be provided as soon as possible."
                            except:
                                datestr = False
                                status = "We could not locate the shipment details for this tracking number. Details are " \
                                         "only available for shipments made within the last 120 days. Please check your " \
                                         "information. "
                        elif shipper == 'FedEx':
                            payload = json.dumps({
                                "trackingInfo": [
                                    {
                                        "trackingNumberInfo": {
                                            "trackingNumber": inquiry_number
                                        }
                                    }
                                ],
                                "includeDetailedScans": False
                            })
                            headers = {
                                'Content-Type': "application/json",
                                'x-customer-transaction-id': tkt,
                                'X-locale': "en_US",
                                'Authorization': f"Bearer {fedex_access_token}",
                            }

                            response = requests.post("https://apis.fedex.com/track/v1/trackingnumbers", data=payload,
                                                     headers=headers)
                            try:
                                # print(json.loads(response.content)['output']['completeTrackResults'][0]['trackResults'][0]['latestStatusDetail'])
                                status = \
                                    json.loads(response.content)['output']['completeTrackResults'][0]['trackResults'][
                                        0][
                                        'latestStatusDetail']['description']
                            except:
                                status = ""
                            real_date = False
                            try:
                                del_date = \
                                    json.loads(response.content)['output']['completeTrackResults'][0]['trackResults'][
                                        0][
                                        'dateAndTimes'][0][
                                        'dateTime']
                                del_date = del_date.replace("T", " ").split()
                                del_date, del_time = del_date
                                del_date = del_date.replace("-", "")
                                del_date = datetime.datetime.strptime(del_date, '%Y%m%d')
                                real_date = del_date.isoformat()[:-3] + 'Z'
                                del_date = f"{str(calendar.month_name[del_date.month])} {str(del_date.day)}," \
                                           f" {str(del_date.year)}".strip()
                                del_time = del_time.replace("-", " ").split()[0]
                                del_time = del_time.replace(":", " ").split()[0:2]
                                delimiter = ":"
                                del_time = delimiter.join(del_time)
                                if del_time == "00:00":
                                    del_time = "EOD"
                                datestr = f"{del_date} by {del_time}"
                            except:
                                datestr = ""
                        else:
                            status = ""
                            datestr = ""
                    else:
                        datestr = ""
                        try:
                            status = sql_row[2].strip()
                            if status.strip() in (
                                    'RDY TO INVOICE', 'RDY TO INV', 'Q', 'Q2', 'INVOICE') and not inquiry_number:
                                status = ("RDY TO INVOICE - No tracking available. This may be a partial shipment. "
                                          "Tracking may be associated with another ticket number.")
                                new_cell = smart.models.Cell()
                                new_cell.column_id = column_map[tracking_status_column]
                                new_cell.value = status
                                new_cell.strict = False
                                get_row = smart.models.Row()
                                get_row.id = row.id
                                get_row.cells.append(new_cell)
                                # Wrap the update_rows call in a lambda to defer execution
                                update_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

                                # Use the retry function to attempt the update
                                updated_row = smartsheet_api_call_with_retry(update_call)
                        except:
                            status = "Ticket not found"
                            if status.strip() == 'Ticket not found' and not inquiry_number:
                                status = "Ticket not found"
                                new_cell = smart.models.Cell()
                                new_cell.column_id = column_map[tracking_status_column]
                                new_cell.value = status
                                new_cell.strict = False
                                get_row = smart.models.Row()
                                get_row.id = row.id
                                get_row.cells.append(new_cell)
                                # Wrap the update_rows call in a lambda to defer execution
                                update_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

                                # Use the retry function to attempt the update
                                updated_row = smartsheet_api_call_with_retry(update_call)
                    logging.warning(f"Ticket: {tkt}")
                    if inquiry_number:
                        logging.warning(f"Tracking number: {inquiry_number}")
                        if tkt:
                            new_cell = smart.models.Cell()
                            new_cell.column_id = column_map[tracking_column]
                            new_cell.value = inquiry_number
                            new_cell.strict = False
                            get_row = smart.models.Row()
                            get_row.id = row.id
                            get_row.cells.append(new_cell)
                            # Wrap the update_rows call in a lambda to defer execution
                            update_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

                            # Use the retry function to attempt the update
                            updated_row = smartsheet_api_call_with_retry(update_call)
                    if status:
                        logging.warning(textwrap.fill(f"Status: {status}", width=(just_len + 2)))
                        if tkt and inquiry_number:
                            if all(substring in status.lower() for substring in ['return', 'sender']):
                                status = "Returned to sender"
                            if get_value(tracking_status_column).lower() != 'delivered':
                                new_cell = smart.models.Cell()
                                new_cell.column_id = column_map[tracking_status_column]
                                new_cell.value = status
                                new_cell.strict = False
                                get_row = smart.models.Row()
                                get_row.id = row.id
                                get_row.cells.append(new_cell)
                                # Wrap the update_rows call in a lambda to defer execution
                                update_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

                                # Use the retry function to attempt the update
                                updated_row = smartsheet_api_call_with_retry(update_call)
                            if ss_date:
                                new_cell = smart.models.Cell()
                                new_cell.column_id = column_map[del_date_column]
                                new_cell.value = ss_date
                                new_cell.strict = False
                                get_row = smart.models.Row()
                                get_row.id = row.id
                                get_row.cells.append(new_cell)
                                # Wrap the update_rows call in a lambda to defer execution
                                update_call = lambda: smart.Sheets.update_rows(sheet_id, [get_row])

                                # Use the retry function to attempt the update
                                updated_row = smartsheet_api_call_with_retry(update_call)
                    del tkt
                    del inquiry_number
                    if "delivered" in status.lower():
                        logging.warning(textwrap.fill(f"Delivery date: {del_date}", width=(just_len + 2)))
                        del status
                        del del_date
                    else:
                        if datestr:
                            logging.warning(textwrap.fill(f"Estimated delivery: {datestr}", width=(just_len + 2)))
                            del datestr
                        else:
                            pass
                    logging.warning("-" * (just_len + 2))
                else:
                    pass


phone_dict = {
    "HT812": "HT812 ATA",
    "HT814": "HT814 ATA",
    "HT818": "HT818 ATA",
    "2200-49530-001": "OBi300",
    "2200-49532-001": "OBi302",
    "2200-49550-001": "OBi504",
    "2200-49552-001": "OBi508",
    "2200-49235-001": "D230",
    "2200-49230-001": "D230 IP Base Kit",
    "GRP2614": "GRP2614 - VoIP phone",
    "GRP2615": "GRP2615 - VoIP phone",
    "GRP2616": "GRP2616 - VoIP phone",
    "2200-66800-001": "IP 8300",
    "2200-66700-001": "IP 8500",
    "2200-66070-001": "IP 8800",
    "KX-TGP600G": "KX Base",
    "KX-TGP600": "KX Bundle",
    "KX-TPA65": "KX Desktop",
    "KX-TPA60": "KX Handset",
    "KX-UDT131": "KX Rugged",
    "KX-UDT121": "KX Slim",
    "KX-TGC352B": "KX-TGC352B",
    "2200-86240-019": "Trio C60",
    "2200-18061-025": "VVX 1500",
    "2200-48820-001": "VVX 250",
    "2200-46135-001": "VVX 300",
    "2200-46135-025": "VVX 300",
    "2200-48300-001": "VVX 301",
    "2200-48300-025": "VVX 301",
    "2200-46161-025": "VVX 310",
    "2200-48350-001": "VVX 311",
    "2200-48350-025": "VVX 311",
    "G2200-48350-025": "VVX 311/TAA",
    "2200-48830-001": "VVX 350",
    "2200-46157-001": "VVX 400",
    "2200-46157-025": "VVX 400",
    "2200-46157-025-VQMON": "VVX 400",
    "2200-48400-001": "VVX 401",
    "2200-48400-025": "VVX 401",
    "2200-46162-025": "VVX 410",
    "2200-48450-001": "VVX 411",
    "2200-48450-025": "VVX 411",
    "2200-48840-001": "VVX 450",
    "2200-48842-025": "VVX 450 OBI",
    "2200-44500-001": "VVX 500",
    "2200-44500-025": "VVX 500",
    "2200-48500-001": "VVX 501",
    "2200-48500-025": "VVX 501",
    "2200-44600-018": "VVX 600",
    "2200-44600-025": "VVX 600",
    "2200-48600-001": "VVX 601",
    "2200-48600-025": "VVX 601",
    "2200-48822-001": "VVX250 OBI",
    "2200-48832-001": "VVX350 OBI",
    "2200-48842-001": "VVX450 OBI",
    "2200-86850-001": "Rove 30",
    "2200-88080-001": "Rove 20",
    "2200-86810-001": "Rove 40",
    "WP820": "WP820 WiFi Phone",
    "8301": "Algo",
    "8180G2": "Algo 81",
    "SN5600/4B/EUI": "Patton 56",
    "SN5301/4B/EUI": "Patton 53",
    "SN5501/4B/EUI": "Patton 55"
}
devices = ""
for i in phone_dict.keys():
    devices += f"\'{i}\', "
devices = devices[:-2]


def dummy_api_call():
    # Simulate an API error
    raise smartsheet.exceptions.ApiError("Simulated API Error")


if __name__ == "__main__":
    print(normalize_ticket_number('CW - 3615596'))
    print(normalize_ticket_number('CW3615596-1'))
    print(normalize_ticket_number('3615596-1'))
    print(normalize_ticket_number('3615596'))

    # UpdateTicketData()
    # ClearHelpers()
    # TrackingUpdate(sheet_id=4591122715568004, ticket_column='Equipment Ticket/PWO',
    #                tracking_column='Tracking Number v2', tracking_status_column='UPS Status v2',
    #                del_date_column='Delivery Date v2')
    # TrackingUpdate(sheet_id=5036994589577092, ticket_column='Equipment Ticket',
    #                tracking_column='Tracking Number', tracking_status_column='Shipping Status',
    #                del_date_column='Delivery Date')
    # # VOIPRouterTemplate()
    # print((lambda s: s[0].upper() + s[1:] + ('.' if not s.endswith('.') else ''))(
    #     requests.get("https://www.affirmations.dev/").json()['affirmation']))
# 164906451
