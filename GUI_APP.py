import sys
import paramiko
import re
import json
import time
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QTextEdit,
    QGroupBox, QTabWidget, QScrollArea, QProgressBar,
    QTextBrowser, QFrame, QSizePolicy, QMessageBox, QSpacerItem,
    QTableWidget, QTableWidgetItem, QHeaderView, QSplitter
)
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QDate, QThread, QTimer
from PyQt5.QtGui import QFont, QTextCursor, QColor
from PyQt5.QtWidgets import QDateEdit
import os
import logging
from logging.handlers import RotatingFileHandler
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import datetime
from functools import wraps

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")


def setup_logging():
    path = "C:/tmp"

    # Check if the directory exists
    if not os.path.exists(path):
        # Create the directory
        os.makedirs(path)
        print(f"Directory {path} created.")
    else:
        print(f"Directory {path} already exists.")

    LOG_FILE = os.path.join(path, f'test_station_interface_{timestamp}.log')

    # Custom formatter that includes function name
    class ContextFormatter(logging.Formatter):
        def format(self, record):
            # Only add function name if it's not already there
            if not hasattr(record, 'func_name'):
                record.func_name = "Internal_Function_driver"
            return super().format(record)

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # File handler
    file_handler = RotatingFileHandler(
        LOG_FILE,
        maxBytes=1024 * 1024,
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # Formatter with function name
    formatter = ContextFormatter(
        '%(asctime)s - %(name)s - %(levelname)s - [%(func_name)s] - %(message)s'
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    # Decorator to add function name to log records
    def log_function(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = logging.getLogger(func.__module__)
            logger.debug(f"Entering {func.__name__}", extra={'func_name': func.__name__})
            try:
                result = func(*args, **kwargs)
                logger.debug(f"Exiting {func.__name__}", extra={'func_name': func.__name__})
                return result
            except Exception as e:
                logger.error(f"Error in {func.__name__}: {str(e)}",
                             exc_info=True,
                             extra={'func_name': func.__name__})
                raise

        return wrapper

    return logger, log_function


# Initialize logging
logger, log_function = setup_logging()


class ExcelLogger:
    def __init__(self, file_path=os.path.join('C:/tmp', f'test_station_records_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')):
        self.file_path = file_path
        self.workbook = None
        self.unit_sheet = None
        self.interlock_sheet = None
        self.self_test_sheet = None
        self.BNC_sheet = None
        self.resistance_sheet = None
        self.Imp_sheet = None
        self.logger1 = logger.getChild('ExcelLogger')
        self.pn=''
        self.sn=''
        self.excel_time=datetime.now().strftime("%Y%m%d_%H%M%S")

        # Create the directory if it doesn't exist
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        # Initialize or load the workbook
        self._init_workbook()

    @log_function
    def _init_workbook(self):
        """Initialize or load the workbook"""
        if os.path.exists(self.file_path):
            try:
                self.workbook = load_workbook(self.file_path)
            except Exception as e:
                self.logger1.error(f"Error loading existing workbook:", exc_info=True,
                                   extra={'func_name': 'init_workbook'})
                self._create_new_workbook()
        else:
            self._create_new_workbook()

    @log_function
    def _create_new_workbook(self):
        """Create a new empty workbook"""
        self.workbook = Workbook()

        # Instead of removing the default sheet, just rename it to one of your expected sheets
        default_sheet = self.workbook.active
        default_sheet.title = "Unit Setup"  # Or whichever sheet you expect to use first

        # Create headers for the default sheet
        self._create_unit_headers(default_sheet)

        try:
            self.workbook.save(self.file_path)
            self.logger1.info(f"Created new Excel file at {self.file_path}",
                              extra={'func_name': 'create_new_workbook'})
        except Exception as e:
            self.logger1.error(f"Error saving new workbook: {e}", exc_info=True,
                               extra={'func_name': 'create_new_workbook'})

    @staticmethod
    def _freq_to_sheet_suffix(freq_text):
        """Convert frequency text to a safe sheet name suffix e.g. '60 MHz' -> '60MHz'"""
        return freq_text.replace(' ', '').replace('.', '_')

    @log_function
    def reset_sheet(self, sheet_name):
        """Clear all data from a specific sheet (except headers) and recreate headers if needed"""
        try:
            if sheet_name not in self.workbook.sheetnames:
                self.logger1.warning(f"Sheet '{sheet_name}' does not exist",
                                     extra={'func_name': 'reset_sheet'})
                return False

            sheet = self.workbook[sheet_name]

            # "Unit Setup" has no header row, so clear from row 1;
            # all other sheets keep the header in row 1.
            if sheet_name == "Unit Setup":
                if sheet.max_row >= 1:
                    sheet.delete_rows(1, sheet.max_row)
            elif sheet.max_row > 1:
                sheet.delete_rows(2, sheet.max_row)  # Delete from row 2 to end

            # Reapply headers based on sheet type
            if sheet_name == "Interlock Test":
                self._create_interlock_headers(sheet)
            elif sheet_name == "Self Test":
                self._create_self_test_headers(sheet)
            elif sheet_name == "Zone1-Inner_Res_scan":
                self._create_resistance_headers(sheet)
            elif sheet_name == "Zone2-Mid_Inner_Res_scan":
                self._create_resistance_headers(sheet)
            elif sheet_name == "Zone3-Mid_Edge_Res_scan":
                self._create_resistance_headers(sheet)
            elif sheet_name == "Zone4-Edge_Res_scan":
                self._create_resistance_headers(sheet)
            elif sheet_name == "Zone5-Outer_Res_scan":
                self._create_resistance_headers(sheet)
            elif "_Imp_scan" in sheet_name:
                self._create_impedance_headers(sheet)
            elif sheet_name == "BNC Port Verification":
                self._create_BNC_headers(sheet)
            elif sheet_name == "Unit Setup":
                self._create_unit_headers(sheet)

            self.workbook.save(self.file_path)
            self.logger1.info(f"Reset sheet '{sheet_name}' successfully",
                              extra={'func_name': 'reset_sheet'})
            return True
        except Exception as e:
            self.logger1.error(f"Error resetting sheet '{sheet_name}': {e}", exc_info=True,
                               extra={'func_name': 'reset_sheet'})
            return False

    def _ensure_sheet_exists(self, sheet_name, create_headers_func):
        """Ensure a sheet exists, creating it if necessary"""
        if sheet_name not in self.workbook.sheetnames:
            sheet = self.workbook.create_sheet(sheet_name)
            create_headers_func(sheet)
            self.workbook.save(self.file_path)
            return sheet
        return self.workbook[sheet_name]

    def _create_unit_headers(self, sheet):
        """Create headers for the unit setup sheet"""
        sheet.column_dimensions['A'].width = 20
        sheet.column_dimensions['B'].width = 30

    def _create_interlock_headers(self, sheet):
        """Create headers for the interlock test sheet"""
        sheet.column_dimensions['A'].width = 20  # Timestamp
        sheet.column_dimensions['B'].width = 20  # Test Name
        sheet.column_dimensions['C'].width = 15  # Open Count
        sheet.column_dimensions['D'].width = 10  # Closed Count
        sheet.column_dimensions['E'].width = 15  # Test Result
        sheet.column_dimensions['F'].width = 30  # Notes

        headers = [
            "Timestamp", "Test Name", "Open Count", "Closed Count",
            "Test Result", "Notes"
        ]

        for col_num, header in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        sheet.freeze_panes = "A2"

    def _create_self_test_headers(self, sheet):
        """Create headers for the self test sheet"""
        sheet.column_dimensions['A'].width = 20  # Timestamp
        sheet.column_dimensions['B'].width = 20  # Unit Identifier
        sheet.column_dimensions['C'].width = 15  # Test Result
        sheet.column_dimensions['D'].width = 30  # Test Details
        sheet.column_dimensions['E'].width = 30  # Notes

        headers = [
            "Timestamp", "Unit Identifier", "Test Result",
            "Test Details", "Notes"
        ]

        for col_num, header in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        sheet.freeze_panes = "A2"

    def _create_resistance_headers(self, sheet):
        """Create headers for the resistance measurements sheet"""
        sheet.column_dimensions['A'].width = 20  # Timestamp
        sheet.column_dimensions['B'].width = 20  # Zone
        sheet.column_dimensions['C'].width = 15  # Setpoint
        sheet.column_dimensions['D'].width = 15  # Resistance
        sheet.column_dimensions['E'].width = 10  # Status
        sheet.column_dimensions['F'].width = 10  # Table Row

        headers = [
            "Timestamp", "Zone", "Setpoint", "Resistance (Ω)",
            "Status", "Table Row"
        ]

        for col_num, header in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        sheet.freeze_panes = "A2"

    def _create_impedance_headers(self, sheet):
        """Create headers for the impedance measurements sheet"""
        sheet.column_dimensions['A'].width = 20  # Timestamp
        sheet.column_dimensions['B'].width = 20  # Zone
        sheet.column_dimensions['C'].width = 20  # Frequency
        sheet.column_dimensions['D'].width = 15  # Setpoint
        sheet.column_dimensions['E'].width = 15  # Real
        sheet.column_dimensions['F'].width = 15  # Image
        sheet.column_dimensions['G'].width = 15  # Impedance
        sheet.column_dimensions['H'].width = 10  # Status

        headers = [
            "Timestamp", "Zone", "Frequency", "Setpoint",
            "Real(Ω)", "Imaginary", "Impedance(Z)", "Status"
        ]

        for col_num, header in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        sheet.freeze_panes = "A2"

    def _create_BNC_headers(self, sheet):
        """Create headers for the BNC Port Verification sheet"""
        sheet.column_dimensions['A'].width = 20  # Timestamp
        sheet.column_dimensions['B'].width = 20  # Zone NAME
        sheet.column_dimensions['C'].width = 20  # TEST_VALUE
        sheet.column_dimensions['D'].width = 15  # STATUS

        headers = [
            "Timestamp", "Zone", "Value(db)", "Status"
        ]

        for col_num, header in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        sheet.freeze_panes = "A2"

    def _create_summary_headers(self, sheet):
            """Create headers for the Summary sheet with additional metadata fields"""
            # Set up metadata fields in rows 1-9
            metadata_fields = [
                ("EID", ""),
                ("SERIAL NUMBER", ""),
                ("MODEL NUMBER", ""),
                ("VERSION", ""),
                ("TESTER NAME", ""),
                ("COMMENT", ""),
                ("START TIME", ""),
                ("END TIME", ""),
                ("OVERALL RESULT", ""),
                ("TEST FIXTURE SN", ""),
                ("VNA SN", ""),
                ("ECAL SN", "")
            ]

            # Write metadata fields with formatting
            for row_num, (field, _) in enumerate(metadata_fields, start=1):
                # Field name cell
                sheet.cell(row=row_num, column=1, value=field)
                sheet.cell(row=row_num, column=1).font = Font(bold=True)

                # Value cell (empty initially)
                sheet.cell(row=row_num, column=2, value="")

                # Format for OVERALL RESULT
                if field == "OVERALL RESULT":
                    sheet.cell(row=row_num, column=2).font = Font(bold=True)
                    sheet.cell(row=row_num, column=2).alignment = Alignment(horizontal='center')

            # Add space between metadata and TESTSTEP section
            sheet.row_dimensions[13].height = 15

            # TESTSTEP section (now starting at row 14)
            sheet.cell(row=14, column=1, value="TESTSTEP").font = Font(bold=True, color="FFFFFF")
            sheet.cell(row=14, column=1).fill = PatternFill(
                start_color="404040", end_color="404040", fill_type="solid")

            sheet.cell(row=14, column=2, value="STATUS").font = Font(bold=True, color="FFFFFF")
            sheet.cell(row=14, column=2).fill = PatternFill(
                start_color="404040", end_color="404040", fill_type="solid")

            # Add space between sections
            sheet.row_dimensions[21].height = 15

            # STEP section headers (now starting at row 22 from column A)
            headers_row22 = ["Step", "Unit", "Low_Limit", "Measure", "High_Limit",
                             "TestStep", "TestPoints", "Status"]
            for col_num, header in enumerate(headers_row22, start=1):  # Starting at column A (1)
                cell = sheet.cell(row=22, column=col_num, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')

            # Set column widths
            sheet.column_dimensions['A'].width = 20  # Step
            sheet.column_dimensions['B'].width = 15  # Unit
            sheet.column_dimensions['C'].width = 15  # Low_Limit
            sheet.column_dimensions['D'].width = 15  # Measure
            sheet.column_dimensions['E'].width = 15  # High_Limit
            sheet.column_dimensions['F'].width = 15  # TestStep
            sheet.column_dimensions['G'].width = 15  # TestPoints
            sheet.column_dimensions['H'].width = 15  # Status

            sheet.freeze_panes = "A23"  # Freeze above the STEP section

    @log_function
    def update_overall_result(self, result, PN='NA', SN='NA'):
        """Update the overall result and rename file accordingly"""
        try:
            result = result.upper()
            if result not in ['PASS', 'FAIL']:
                self.logger1.warning(f"Invalid result: {result}. Must be 'PASS' or 'FAIL'")
                return False

            if PN != 'NA' and SN != 'NA':
                self.pn = PN
                self.sn = SN

            # Create new filename based on result
            new_filename = f"{self.pn}_{self.sn}_{self.excel_time}_{result}.xlsx"
            new_file_path = os.path.join("C:\\tmp", new_filename)
            print("new_fil", new_file_path)

            # If file already exists with different name, rename it
            if self.file_path != new_file_path:
                if os.path.exists(self.file_path):
                    # Close the workbook before renaming
                    self.workbook.close()

                    # Rename the file
                    os.rename(self.file_path, new_file_path)
                    self.file_path = new_file_path

                    # Reopen the workbook
                    self.workbook = load_workbook(self.file_path)
                    self.logger1.info(f"Renamed file to: {new_file_path}",
                                      extra={'func_name': 'update_overall_result'})

            return True

        except Exception as e:
            self.logger1.error(f"Error updating overall result: {e}", exc_info=True,
                               extra={'func_name': 'update_overall_result'})
            return False

    @log_function
    def log_unit_setup(self, unit_data):
        """Log unit setup data to the Excel file"""
        try:
            # Ensure sheet exists
            self.unit_sheet = self._ensure_sheet_exists(
                "Unit Setup",
                self._create_unit_headers
            )

            # Get the next available row
            row_num = 1 if self.unit_sheet.max_row == 1 and all(
                cell.value is None for cell in self.unit_sheet[1]) else self.unit_sheet.max_row + 1

            fields = [
                ("Vendor Name", unit_data.get('Vendor_name', '')),
                ("Fixture Number", unit_data.get('Fixture_number', '')),
                ("Test Operator Name", unit_data.get('test_operator_name', '')),
                ("Test Date", unit_data.get('test_date', '')),
                ("VNA Calibration Date", unit_data.get('vna_calibration_date', '')),
                ("VNA SN", unit_data.get('vna_sn', '')),
                ("Ecal SN", unit_data.get('ecal_sn', '')),
                ("PCB Control Part Number", unit_data.get('pcb_part_number', '')),
                ("PCB Control Revision", unit_data.get('pcb_revision', '')),
                ("PCB Control Serial Number", unit_data.get('pcb_serial_number', '')),
                ("Assembly Part Number", unit_data.get('assembly_part_number', '')),
                ("Assembly Revision", unit_data.get('assembly_revision', '')),
                ("Assembly Serial Number", unit_data.get('assembly_serial_number', '')),
                ("Product ID", unit_data.get('product_id', '')),
                ("ESI Revision", unit_data.get('esi_revision', '')),
                ("Configuration ID", unit_data.get('configuration_id', '')),
                ("EtherCAT Address", unit_data.get('ethercat_address', '')),
                ("Firmware Version", unit_data.get('firmware_version', ''))
            ]

            for i, (field_name, field_value) in enumerate(fields, start=0):
                field_row = row_num + i
                self.unit_sheet.cell(row=field_row, column=1, value=field_name)
                self.unit_sheet.cell(row=field_row, column=1).font = Font(bold=True)
                self.unit_sheet.cell(row=field_row, column=2, value=field_value)
                for col in [1, 2]:
                    self.unit_sheet.cell(row=field_row, column=col).alignment = Alignment(
                        horizontal='left', vertical='center'
                    )

            self.unit_sheet.append([])
            self.workbook.save(self.file_path)
            self.logger1.info(f"Logged unit setup data to {self.file_path}",
                              extra={'func_name': 'log_unit_setup'})
            return True
        except Exception as e:
            self.logger1.error(f"Error logging unit setup data: {e}", exc_info=True,
                               extra={'func_name': 'log_unit_setup'})
            return False

    @log_function
    def log_interlock_test(self, test_name, test_passed, open_count, closed_count, notes=""):
        """Log interlock test results to the Excel file"""
        try:
            # Ensure sheet exists
            self.interlock_sheet = self._ensure_sheet_exists(
                "Interlock Test",
                self._create_interlock_headers
            )

            # Get the next available row
            row_num = self.interlock_sheet.max_row + 1

            # Write data with timestamp
            self.interlock_sheet.cell(
                row=row_num, column=1,
                value=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
            self.interlock_sheet.cell(row=row_num, column=2, value=test_name)

            # Test Result with color coding
            result_cell = self.interlock_sheet.cell(
                row=row_num, column=5,
                value="PASS" if test_passed else "FAIL"
            )
            result_cell.font = Font(bold=True)
            result_cell.fill = PatternFill(
                start_color="00AA00" if test_passed else "FF0000",
                end_color="00AA00" if test_passed else "FF0000",
                fill_type="solid"
            )
            result_cell.alignment = Alignment(horizontal='center')

            self.interlock_sheet.cell(row=row_num, column=3, value=open_count)
            self.interlock_sheet.cell(row=row_num, column=4, value=closed_count)
            self.interlock_sheet.cell(row=row_num, column=6, value=notes)

            # Save the workbook
            self.workbook.save(self.file_path)
            self.logger1.info(f"Logged interlock test result to {self.file_path}",
                              extra={'func_name': 'Log_interlock_test'})
            return True
        except Exception as e:
            self.logger1.error(f"Error logging interlock test result: {e}", exc_info=True,
                               extra={'func_name': 'Log_interlock_test'})
            return False

    @log_function
    def log_self_test(self, unit_identifier, test_passed, test_details="", notes=""):
        """Log self test results to the Excel file"""
        try:
            # Ensure sheet exists
            self.self_test_sheet = self._ensure_sheet_exists(
                "Self Test",
                self._create_self_test_headers
            )

            # Get the next available row
            row_num = self.self_test_sheet.max_row + 1

            # Write data with timestamp
            self.self_test_sheet.cell(
                row=row_num, column=1,
                value=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
            self.self_test_sheet.cell(row=row_num, column=2, value=unit_identifier)

            # Test Result with color coding
            result_cell = self.self_test_sheet.cell(
                row=row_num, column=3,
                value="PASS" if test_passed else "FAIL"
            )
            result_cell.font = Font(bold=True)
            result_cell.fill = PatternFill(
                start_color="00AA00" if test_passed else "FF0000",
                end_color="00AA00" if test_passed else "FF0000",
                fill_type="solid"
            )
            result_cell.alignment = Alignment(horizontal='center')

            self.self_test_sheet.cell(row=row_num, column=4, value=test_details)
            self.self_test_sheet.cell(row=row_num, column=5, value=notes)

            # Save the workbook
            self.workbook.save(self.file_path)
            self.logger1.info(f"Logged self test result to {self.file_path}",
                              extra={'func_name': 'Log_self_test'})
            return True
        except Exception as e:
            self.logger1.error(f"Error logging self test result: {e}", exc_info=True,
                               extra={'func_name': 'Log_self_test'})
            return False

    @log_function
    def log_resistance_measurement(self, measurement_data, sheet_name):
        """Log resistance measurement to combined sheet"""
        try:
            # Ensure sheet exists
            self.resistance_sheet = self._ensure_sheet_exists(
                sheet_name,
                self._create_resistance_headers
            )

            # Write measurement data
            row_num = self.resistance_sheet.max_row + 1
            self.resistance_sheet.cell(row=row_num, column=1, value=measurement_data['timestamp'])
            self.resistance_sheet.cell(row=row_num, column=2, value=measurement_data['zone_title'])
            self.resistance_sheet.cell(row=row_num, column=3, value=measurement_data['setpoint'])
            self.resistance_sheet.cell(row=row_num, column=4, value=measurement_data['resistance'])

            # Status with color coding
            status_cell = self.resistance_sheet.cell(
                row=row_num, column=5,
                value=measurement_data['status']
            )
            if measurement_data['status'] == "PASS":
                status_cell.fill = PatternFill(
                    start_color="00AA00",  # Green
                    fill_type="solid"
                )
            else:
                status_cell.fill = PatternFill(
                    start_color="FF0000",  # Red
                    fill_type="solid"
                )

            self.resistance_sheet.cell(row=row_num, column=6, value=measurement_data['table_row'])

            # Save workbook
            self.workbook.save(self.file_path)
            self.logger1.info(
                f"Logged resistance data to combined worksheet",
                extra={'func_name': 'log_resistance_measurement'}
            )
            return True
        except Exception as e:
            self.logger1.error(
                f"Error logging resistance data: {str(e)}",
                exc_info=True,
                extra={'func_name': 'log_resistance_measurement'}
            )
            return False


    @log_function
    def log_summary(self, metadata=None, teststep_data=None, step_data=None, update_existing=True):
        """Log data to the Summary sheet with option to update existing entries"""
        try:
            # Ensure sheet exists
            summary_sheet = self._ensure_sheet_exists(
                "Summary",
                self._create_summary_headers
            )

            # Update metadata if provided
            if metadata:
                metadata_mapping = {
                    'eid': 1,
                    'serial_number': 2,
                    'model_number': 3,
                    'version': 4,
                    'tester_name': 5,
                    'comment': 6,
                    'start_time': 7,
                    'end_time': 8,
                    'overall_result': 9,
                    'test_fixture_sn': 10,
                    'vna_sn': 11,
                    'ecal_sn': 12
                }

                for key, value in metadata.items():
                    if key.lower() in metadata_mapping:
                        row_num = metadata_mapping[key.lower()]
                        summary_sheet.cell(row=row_num, column=2, value=value)

                        if key.lower() == 'overall_result':
                            result_cell = summary_sheet.cell(row=row_num, column=2)
                            result_cell.font = Font(bold=True)
                            result_cell.alignment = Alignment(horizontal='center')
                            if str(value).upper() == "PASS":
                                result_cell.fill = PatternFill(
                                    start_color="00AA00", end_color="00AA00", fill_type="solid")
                                result_cell.font = Font(color="FFFFFF", bold=True)
                            elif str(value).upper() == "FAIL":
                                result_cell.fill = PatternFill(
                                    start_color="FF0000", end_color="FF0000", fill_type="solid")
                                result_cell.font = Font(color="FFFFFF", bold=True)

            # Update TESTSTEP data directly if provided
            if teststep_data:
                teststep_name = teststep_data.get('teststep', '')
                teststep_status = teststep_data.get('status', '')

                if teststep_name:
                    # Find existing teststep in TESTSTEP section (rows 15-20, column 1)
                    teststep_updated = False
                    for row in range(15, 21):  # TESTSTEP section rows
                        existing_teststep = summary_sheet.cell(row=row, column=1).value
                        if existing_teststep == teststep_name:
                            # Update existing teststep
                            status_cell = summary_sheet.cell(row=row, column=2, value=teststep_status)
                            self._apply_status_formatting(status_cell, teststep_status)
                            teststep_updated = True
                            break

                    # If not found and we have space, add to first empty row
                    if not teststep_updated:
                        for row in range(15, 21):
                            existing_teststep = summary_sheet.cell(row=row, column=1).value
                            if not existing_teststep or existing_teststep == "":
                                # Add new teststep
                                summary_sheet.cell(row=row, column=1, value=teststep_name)
                                status_cell = summary_sheet.cell(row=row, column=2, value=teststep_status)
                                self._apply_status_formatting(status_cell, teststep_status)
                                teststep_updated = True
                                break

            # Log STEP data if provided
            if step_data:
                testpoints = step_data.get('testpoints', '')

                # Find existing row if update_existing is True - search in TestPoints column (column G, index 7)
                step_row = None
                if update_existing and testpoints:
                    # Search from row 23 onwards for matching testpoints
                    for row in range(23, summary_sheet.max_row + 1):
                        cell_value = summary_sheet.cell(row=row, column=7).value  # Column G (7) is TestPoints
                        if cell_value == testpoints:
                            step_row = row
                            break

                # If not found or not updating, use next available row
                if step_row is None:
                    step_row = 23
                    # Find first empty row in the STEP section
                    while (summary_sheet.cell(row=step_row, column=7).value is not None and
                           summary_sheet.cell(row=step_row, column=7).value != ""):
                        step_row += 1

                    # For new rows, write step data including the step column
                    summary_sheet.cell(row=step_row, column=1, value=step_data.get('step', ''))  # Column A: Step
                # For existing rows, DO NOT update the Step column (Column A) - keep it as is

                # Update columns B to H (Unit to Status) - preserving Step column (A)
                summary_sheet.cell(row=step_row, column=2, value=step_data.get('unit', ''))  # Column B: Unit
                summary_sheet.cell(row=step_row, column=3, value=step_data.get('low_limit', ''))  # Column C: Low_Limit
                summary_sheet.cell(row=step_row, column=4, value=step_data.get('measure', ''))  # Column D: Measure
                summary_sheet.cell(row=step_row, column=5,
                                   value=step_data.get('high_limit', ''))  # Column E: High_Limit
                summary_sheet.cell(row=step_row, column=6, value=step_data.get('teststep', ''))  # Column F: TestStep
                summary_sheet.cell(row=step_row, column=7, value=testpoints)  # Column G: TestPoints



                # Write STATUS with color coding in column H (index 8)
                status = step_data.get('status', '')
                status_cell = summary_sheet.cell(row=step_row, column=8, value=status)  # Column H: Status
                self._apply_status_formatting(status_cell, status)

            # If step_data was provided, update TESTSTEP section automatically based on STEP section data
            # BUT only update teststeps that are managed by STEP data (not manually set ones)
            if step_data:
                self._update_teststep_from_step_data_preserve_manual(summary_sheet)

            # AUTOMATICALLY UPDATE OVERALL RESULT BASED ON TESTSTEP STATUS (ROWS 15-20, COLUMN B)
            self._update_overall_result_based_on_teststep_status(summary_sheet)

            # Save the workbook
            self.workbook.save(self.file_path)
            self.logger1.info("Logged summary data successfully",
                              extra={'func_name': 'log_summary'})
            return True

        except Exception as e:
            self.logger1.error(f"Error logging summary data: {e}", exc_info=True,
                               extra={'func_name': 'log_summary'})
            return False

    @log_function
    def _update_teststep_from_step_data_preserve_manual(self, summary_sheet):
        """Update TESTSTEP section based on STEP section data but preserve manually set teststeps"""
        try:
            # Get current manually set teststeps (rows 15-20)
            manual_teststeps = {}
            for row in range(15, 21):
                teststep_name = summary_sheet.cell(row=row, column=1).value
                if teststep_name:
                    manual_teststeps[teststep_name] = {
                        'row': row,
                        'status': summary_sheet.cell(row=row, column=2).value
                    }

            # Dictionary to store teststep statuses from STEP data
            step_teststep_status_map = {}

            # Collect all teststeps and their statuses from STEP section (column F and H)
            for row in range(23, summary_sheet.max_row + 1):
                teststep_cell = summary_sheet.cell(row=row, column=6)  # Column F: TestStep
                status_cell = summary_sheet.cell(row=row, column=8)  # Column H: Status

                teststep_name = teststep_cell.value
                status_value = str(status_cell.value).upper().strip() if status_cell.value else ""

                # Only process teststeps that are NOT manually managed
                if teststep_name and teststep_name not in manual_teststeps:
                    if teststep_name not in step_teststep_status_map:
                        step_teststep_status_map[teststep_name] = "PASS"  # Start with PASS assumption

                    # If any step in a teststep fails, the entire teststep fails
                    if status_value == "FAIL":
                        step_teststep_status_map[teststep_name] = "FAIL"

            # Update TESTSTEP section (rows 15-20) - only for step-managed teststeps
            current_teststep_row = 15

            # First, keep existing manual teststeps
            for row in range(15, 21):
                teststep_name = summary_sheet.cell(row=row, column=1).value
                if teststep_name and teststep_name in manual_teststeps:
                    # Keep manual teststep as is
                    current_teststep_row += 1

            # Then add step-managed teststeps to remaining rows
            for teststep_name, status in step_teststep_status_map.items():
                if current_teststep_row > 20:  # Don't exceed the TESTSTEP section
                    break

                # Write teststep name
                summary_sheet.cell(row=current_teststep_row, column=1, value=teststep_name)

                # Write status with color coding
                status_cell = summary_sheet.cell(row=current_teststep_row, column=2, value=status)
                self._apply_status_formatting(status_cell, status)

                current_teststep_row += 1

            # Clear any remaining rows in TESTSTEP section
            for row in range(current_teststep_row, 21):
                # Only clear if not a manual teststep
                existing_teststep = summary_sheet.cell(row=row, column=1).value
                if existing_teststep not in manual_teststeps:
                    summary_sheet.cell(row=row, column=1, value="")
                    summary_sheet.cell(row=row, column=2, value="")

            self.logger1.info(
                f"Updated TESTSTEP section - Manual: {len(manual_teststeps)}, Step-managed: {len(step_teststep_status_map)}",
                extra={'func_name': '_update_teststep_from_step_data_preserve_manual'})

        except Exception as e:
            self.logger1.error(f"Error updating TESTSTEP from STEP data: {e}", exc_info=True,
                               extra={'func_name': '_update_teststep_from_step_data_preserve_manual'})

    @log_function
    def _apply_status_formatting(self, cell, status):
        """Apply consistent status formatting to a cell"""
        status = str(status).upper().strip() if status else ""
        cell.font = Font(bold=True)

        if status == "PASS":
            cell.fill = PatternFill(start_color="00AA00", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)
        elif status == "FAIL":
            cell.fill = PatternFill(start_color="FF0000", fill_type="solid")
            cell.font = Font(color="FFFFFF", bold=True)

    @log_function
    def _update_overall_result_based_on_teststep_status(self, summary_sheet):
        """Update overall result based only on TESTSTEP status (rows 15-20, column B)"""
        try:
            overall_result = "PASS"  # Start with PASS assumption

            # Check TESTSTEP status from rows 15-20, column B
            for row in range(15, 21):  # TESTSTEP section rows 15-20
                status_cell = summary_sheet.cell(row=row, column=2)  # Column B: Status
                status_value = str(status_cell.value).upper().strip() if status_cell.value else ""

                # If any teststep status is FAIL or empty, overall result becomes FAIL
                if status_value == "FAIL" or status_value == "":
                    overall_result = "FAIL"
                    break  # No need to check further if we found one FAIL

            # Update overall result in row 9, column 2
            result_cell = summary_sheet.cell(row=9, column=2, value=overall_result)
            result_cell.font = Font(bold=True)
            result_cell.alignment = Alignment(horizontal='center')
            self._apply_status_formatting(result_cell, overall_result)

            self.logger1.info(f"Updated overall result to: {overall_result} (based on TESTSTEP status)",
                              extra={'func_name': '_update_overall_result_based_on_teststep_status'})

        except Exception as e:
            self.logger1.error(f"Error updating overall result from teststep status: {e}", exc_info=True,
                               extra={'func_name': '_update_overall_result_based_on_teststep_status'})

    @log_function
    def log_Imp_measurement(self, measurement_data, sheet_name):
        """Log impedance measurement to combined sheet"""
        try:
            # Ensure sheet exists
            self.Imp_sheet = self._ensure_sheet_exists(
                sheet_name,
                self._create_impedance_headers
            )

            row_num = self.Imp_sheet.max_row + 1
            self.Imp_sheet.cell(row=row_num, column=1, value=measurement_data['timestamp'])
            self.Imp_sheet.cell(row=row_num, column=2, value=measurement_data['zone_title'])
            self.Imp_sheet.cell(row=row_num, column=3, value=measurement_data['Frequency'])
            self.Imp_sheet.cell(row=row_num, column=4, value=measurement_data['setpoint'])
            self.Imp_sheet.cell(row=row_num, column=5, value=measurement_data['Real'])
            self.Imp_sheet.cell(row=row_num, column=6, value=measurement_data['Imag'])
            self.Imp_sheet.cell(row=row_num, column=7, value=measurement_data['Z'])
            self.Imp_sheet.cell(row=row_num, column=8, value=measurement_data['status'])

            # Status with color coding
            status_cell = self.Imp_sheet.cell(
                row=row_num, column=8,
                value=measurement_data['status']
            )
            if measurement_data['status'] == "PASS":
                status_cell.fill = PatternFill(
                    start_color="00AA00",  # Green
                    fill_type="solid"
                )
            else:
                status_cell.fill = PatternFill(
                    start_color="FF0000",  # Red
                    fill_type="solid"
                )

            # Save workbook
            self.workbook.save(self.file_path)
            self.logger1.info(
                f"Logged impedance data to combined worksheet",
                extra={'func_name': 'log_impedance_measurement'}
            )
            return True
        except Exception as e:
            self.logger1.error(
                f"Error logging impedance data: {str(e)}",
                exc_info=True,
                extra={'func_name': 'log_impedance_measurement'}
            )
            return False



    @log_function
    def log_BNC_measurement(self, test_zone, test_details, test_passed):
        """Log BNC measurement to combined sheet"""
        try:
            # Ensure sheet exists
            self.BNC_sheet = self._ensure_sheet_exists(
                "BNC Port Verification",
                self._create_BNC_headers
            )

            # Get the next available row
            row_num = self.BNC_sheet.max_row + 1

            # Write data with timestamp
            self.BNC_sheet.cell(
                row=row_num, column=1,
                value=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
            self.BNC_sheet.cell(row=row_num, column=2, value=test_zone)
            self.BNC_sheet.cell(row=row_num, column=3, value=test_details)
            print(test_passed)
            # Test Result with color coding
            result_cell = self.BNC_sheet.cell(
                row=row_num, column=4,
                value="PASS" if test_passed else "FAIL"
            )
            result_cell.font = Font(bold=True)
            result_cell.fill = PatternFill(
                start_color="00AA00" if test_passed else "FF0000",
                end_color="00AA00" if test_passed else "FF0000",
                fill_type="solid"
            )
            result_cell.alignment = Alignment(horizontal='center')

            # Save the workbook
            self.workbook.save(self.file_path)
            self.logger1.info(f"Logged BNC test result to {self.file_path}",
                              extra={'func_name': 'Log_BNC_test'})
            return True
        except Exception as e:
            self.logger1.error(f"Error logging BNC test result: {e}", exc_info=True,
                               extra={'func_name': 'Log_BNC_test'})
            return False


# Initialize the Excel logger
excel_logger = ExcelLogger()


class SSH_setup:
    def __init__(self):
        self.is_connect = False
        self.ssh = None
        #self.host = "192.168.1.2"
        self.host = "10.119.9.225"
        self.port = 22
        self.username = "robot"
        self.password = "robot"
        self.script_path = "/home/robot/Manufacturing_test/aipc_beta/test.py ecat"
        self.timeout = 10  # seconds
        #self.config = self.load_config()
        self.logger2 = logger.getChild('SSH_setup')


    @log_function
    def Connect_RPI(self, host=None, port=None, username=None, password=None):
        try:
            # Use provided parameters or fall back to instance variables
            host = host or self.host
            port = port or self.port
            username = username or self.username
            password = password or self.password

            self.ssh = paramiko.SSHClient()
            self.ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            self.ssh.connect(
                host,
                port,
                username,
                password,
                timeout=self.timeout
            )
            self.is_connect = True

            # logger.info("SSH connection established successfully")
            self.logger2.info(f"SSH connection established successfully",
                              extra={'func_name': 'Connect_RPI'})

            return True, "Connected successfully"
        except paramiko.AuthenticationException:
            error_msg = "Authentication failed, please verify your credentials"
            self.logger2.error(f"{error_msg}", exc_info=True,
                               extra={'func_name': 'Connect_RPI'})
            return False, error_msg
        except paramiko.SSHException as e:
            error_msg = f"SSH error: {str(e)}"
            self.logger2.error(f"{error_msg}", exc_info=True,
                               extra={'func_name': 'Connect_RPI'})
            return False, error_msg
        except Exception as e:
            error_msg = f"Connection error: {str(e)}"
            self.logger2.error(f"{error_msg}", exc_info=True,
                               extra={'func_name': 'Connect_RPI'})
            return False, error_msg

    def SSH_com(self, command, script_path=None):
        if not self.is_connect or not self.ssh:
            return "", "Not connected to SSH"


        script_path = script_path or self.script_path

        try:
            stdin, stdout, stderr = self.ssh.exec_command(f'sudo python3 {script_path} {command}')
            stdout_data = stdout.read().decode()
            stderr_data = stderr.read().decode()
            return stdout_data, stderr_data
        except Exception as e:
            return "", f"Command execution failed: {str(e)}"

    def SSH_com_stream(self, script_path, command):
        if not self.is_connect:
            raise Exception("SSH connection not established")

        # Run the Python script
        stdin, stdout, stderr = self.ssh.exec_command(f'sudo python3 {script_path} {command}', get_pty=True)

        # Continuously read the output and error
        while True:
            line = stdout.readline()
            if not line:
                break
            yield line.strip()

    @log_function
    def SSH_disconnect(self):
        try:
            if self.ssh:
                self.ssh.close()
                # logger.info("SSH connection closed")
                self.logger2.info(f"SSH connection closed",
                                  extra={'func_name': 'SSH_disconnect'})

        except Exception as e:
            # logger.error(f"Error disconnecting SSH: {str(e)}")
            self.logger2.error(f"Error disconnecting SSH: {str(e)}", exc_info=True,
                               extra={'func_name': 'SSH_disconnect'})
        finally:
            self.is_connect = False


class Worker(QThread):
    output_ready = pyqtSignal(str)
    finished_signal = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, ssh_handler, script_path, command, timeout=30):
        super().__init__()
        self.ssh_handler = ssh_handler
        self.script_path = script_path
        self.command = command
        self._is_running = True
        self.work_timeout = 30
        self.t = timeout
        self.logger3 = logger.getChild('SSH_setup')

    def run(self):
        try:
            if not self._is_running:
                return

            # Connect SSH
            success, message = self.ssh_handler.Connect_RPI()
            if not success:
                self.error_occurred.emit(f"SSH Connection Failed: {message}")
                return

            # Execute command
            stdin, stdout, stderr = self.ssh_handler.ssh.exec_command(
                f'sudo python3 {self.script_path} {self.command}',
                get_pty=True,
                timeout=self.t
            )

            self.logger3.info(f"Command output received",
                             extra={'func_name': f'sudo python3 {self.script_path} {self.command}'})

            if stdout:
                self.logger3.debug(f"stdout:\n{stdout}",
                                  extra={'func_name': f'sudo python3 {self.script_path} {self.command}'})
            if stderr:
                self.logger3.error(f"stderr:\n{stderr}",
                                  extra={'func_name': f'sudo python3 {self.script_path} {self.command}'})

            while self._is_running:
                line = stdout.readline()
                self.logger3.info(f"{self.script_path} {self.command} {line}")
                if not line and self.command != "dimm":
                    break
                if self._is_running == False:
                    break
                self.output_ready.emit(line.strip())


        except Exception as e:
            self.error_occurred.emit(f"Error during execution: {str(e)}")
        finally:
            self.cleanup()
            self.finished_signal.emit()

    def stop(self):
        self._is_running = False
        self.cleanup()

        # if self.isRunning():
        #    self.terminate()
        #    self.wait(2000)

    def cleanup(self):
        if hasattr(self.ssh_handler, 'SSH_disconnect'):
            self.ssh_handler.SSH_disconnect()


class TestStationInterface(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_logger = excel_logger
        self.setWindowTitle("Test Station Interface")
        self.config = ''
        self.assembly_suffix = None

        # Set minimum size instead of fixed geometry
        self.setMinimumSize(1000, 700)

        # Central widget with layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(8, 8, 8, 8)
        self.main_layout.setSpacing(8)
        self.Zone1_Inner_res = 0
        self.Zone2_Mid_Inner_res = 0
        self.Zone3_Mid_Edge_res = 0
        self.Zone4_Edge_res = 0
        self.Zone5_Outer_res = 0
        self.Zone1_Inner_imp = 0
        self.Zone2_Mid_Inner_imp = 0
        self.Zone3_Mid_Edge_imp = 0
        self.Zone4_Edge_imp = 0
        self.Zone5_Outer_imp = 0
        #self.config_transfer()

        # Initialize other components
        self.ssh_handler = SSH_setup()
        self.console_output = None
        self.check = False
        self.handling_flag = 0
        self.Firmware_check = True
        self.worker = None
        self.worker_thread = None
        self.open_count = 0
        self.closed_count = 0
        self.logger = logger.getChild('TestStationInterface')
        self.check_true = 0
        self.dimm_timer = QTimer()
        self.vna_timer = QTimer()
        self.vna_timer.timeout.connect(self.update_vna_progress)
        self.dimm_timer.timeout.connect(self.update_dimm_progress)
        self.dimm_progress_value = 0
        self.vna_progress_value = 0
        self.names = ''
        self.unit_test = 0
        self.impedance_scan = 0
        self.self_t = 0
        self.Res_scan = 0
        self.bnc_t = 0
        self.VNA_c = 0
        self.DIMM_CAL = 0
        self.interlock_t = 0
        self.step_no = 0
        self.resistance_test = 'PASS'
        self.Impedance_test = 'PASS'
        self.over_all_result = 'PASS'
        self.test_result = 'PASS'
        self.PN = ''
        self.SN = ''
        self.init_ui()
        self.stop_increment = False

    def closeEvent(self, event):
        self.cleanup_resources()
        event.accept()

    """      
    def config_transfer(self):
        #host = "192.168.1.2"  # Replace with your Pi's IP
        host = "10.119.28.136"
        username = "robot"
        password = "robot"  # Default, change if needed

        # File paths
        local_file = "C:\\Config\\config.json"  # Windows path
        remote_file = "/home/robot/Manufacturing_test/aipc_beta/config.json"  # Destination on Pi
        local_file1 = "C:\\Config\\zone_config.cfg"  # Windows path
        remote_file1 = "/home/robot/Manufacturing_test/aipc_beta/zone_config.cfg"  # Destination on Pi

        # Initialize SSH client
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(host, username=username, password=password)

        # Transfer file using SFTP
        sftp = ssh.open_sftp()
        sftp.put(local_file, remote_file)
        sftp.put(local_file1, remote_file1)
        sftp.close()
        ssh.close()
    """

    def config_transfer(self, suffix):
        """
        Transfer config files to RPI based on assembly part number suffix.
        suffix: '003' or '004' - determines which config directory to use
        """
        # host = "192.168.1.2"  # Replace with your Pi's IP
        host = "10.119.9.225"
        username = "robot"
        password = "robot"  # Default, change if needed

        # Determine config directory based on suffix
        if suffix is None:
            config_dir = "C:\\Config"
        else:
            config_dir = f"C:\\Config\\{suffix}"

        # File paths
        local_file = f"{config_dir}\\config.json"  # Windows path
        remote_file = "/home/robot/Manufacturing_test/aipc_beta/config.json"  # Destination on Pi
        local_file1 = f"{config_dir}\\zone_config.cfg"  # Windows path
        remote_file1 = "/home/robot/Manufacturing_test/aipc_beta/zone_config.cfg"  # Destination on Pi

        if not os.path.exists(local_file):
            error_msg = f"Configuration file not found: {local_file}"
            self.append_console_message(error_msg + "\n", is_error=True)
            self.logger.error(error_msg, extra={'func_name': 'config_transfer'})
            QMessageBox.critical(self, "ERROR : ", error_msg)
            return False


        if not os.path.exists(local_file1):
            error_msg = f"Zone configuration file not found: {local_file1}"
            self.append_console_message(error_msg + "\n", is_error=True)
            self.logger.error(error_msg, extra={'func_name': 'config_transfer'})
            QMessageBox.critical(self, "ERROR : ", error_msg)
            return False


        # Initialize SSH client
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(host, username=username, password=password)

        # Transfer file using SFTP
        sftp = ssh.open_sftp()
        sftp.put(local_file, remote_file)
        sftp.put(local_file1, remote_file1)
        sftp.close()
        ssh.close()
        return True

    def _extract_assembly_suffix(self, assembly_part_number):
        """
        Extract the suffix (003 or 004) from assembly part number.
        Format: XXX-AXXXXX-003 or XXX-AXXXXX-004
        Returns: '003', '004', or None if invalid format
        """
        parts = assembly_part_number.split('-')
        if len(parts) == 3:
            suffix = parts[2]
            if suffix in ['003', '004', '005']:
                return suffix
        return None

    def load_config(self, suffix):
        try:
            if suffix is None:
                raise ValueError("Assembly suffix is not set; cannot determine config path.")
            with open(rf'C:\Config\{suffix}\config.json') as f:
                return json.load(f)
        except Exception as e:
            # logger.error(f"Error loading config: {str(e)}")
            self.logger.error(f"Error loading config: {str(e)}", exc_info=True,
                               extra={'func_name': 'load_config'})

            return {"expected_firmware_version": "0.0.0"}

    def create_otp_file(self, fname: str,
                        app_pn: str, app_sn: str,
                        assy_pn: str, assy_sn: str) -> None:
        """
        Creates an OTP (One-Time Program) file with board and assembly information.

        Args:
            fname: File name/path to create
            app_pn: Application board part number
            app_sn: Application board serial number
            assy_pn: Assembly part number
            assy_sn: Assembly serial number

        Raises:
            IOError: If file cannot be written
        """
        try:
            header = "# LAM Research App Board and Assy Part and Serial Number"
            content = [
                header,
                f"Board PN: {app_pn}",
                f"Board SN: {app_sn}",
                f"Assy PN: {assy_pn}",
                f"Assy SN: {assy_sn}",
                ""  # Adds final newline
            ]
            self.append_console_message("Write APPOTP FILE")
            with open(fname, 'w', encoding='utf-8') as f:
                f.write('\n'.join(content))

        except IOError as e:
            self.append_console_message("Failed to create OTP file ",is_error= True)
            raise IOError(f"Failed to create OTP file {fname}") from e

    def start_dimm_progress(self):
        """Start the 12-second progress timer"""
        self.dimm_progress_value = 0
        self.dimm_progress.setValue(0)
        self.dimm_timer.start(130)  # 120ms interval for 12 seconds (100*120ms=12s)

    def start_vna_progress(self):
        """Start the 12-second progress timer"""
        self.vna_progress_value = 0
        self.vna_progress.setValue(0)
        self.vna_timer.start(1800)  # 1800ms interval for 180 seconds (100*1800ms=12s)

    def update_dimm_progress(self):
        """Update progress bar incrementally"""
        self.dimm_progress_value += 1
        self.dimm_progress.setValue(self.dimm_progress_value)

        if self.dimm_progress_value >= 100:
            self.dimm_timer.stop()
            self.dimm_progress.setValue(100)

    def update_vna_progress(self):
        """Update progress bar incrementally"""
        self.vna_progress_value += 1
        self.vna_progress.setValue(self.vna_progress_value)

        if self.vna_progress_value >= 100:
            self.vna_timer.stop()
            self.vna_progress.setValue(100)

    def validate_part_number(self, part_number, part_type):
        pattern = r"^\d{3}-[A-Z]\d{5}-\d{3}$"
        if not re.match(pattern, part_number):
            QMessageBox.critical(
                self,
                "Validation Error",
                f"Invalid {part_type} number format!\n\n"
                f"Entered: {part_number}\n"
                "Expected format: XXX-AXXXXX-XXX\n"
                "(3 digits, hyphen, uppercase letter, 5 digits, hyphen, 3 digits)"
            )
            return False
        return True

    def validate_revision_number(self, serial_number, part_type):
        pattern = r"^[a-zA-Z]$"
        if not re.match(pattern, serial_number):
            QMessageBox.critical(
                self,
                "Validation Error",
                f"Invalid {serial_number} number format!\n\n"
                f"Entered: {serial_number}\n"
                "Expected format: [A-Z]or[a-z]\n"
            )
            return False
        return True

    def parse_ssh_output(self, output):
        results = {
            'product_id': 'Not found',
            'esi_revision': 'Not found',
            'firmware_version': 'Not found',
            'ethercat_address': 'Not found',
            'pcb_part_number': 'Not found',
            'pcb_serial_number': 'Not found',
            'assembly_part_number': 'Not found',
            'assembly_serial_number': 'Not found'
        }

        # Extract standard fields
        product_code_match = re.search(r'Product Code: ([^\n]+)', output, re.DOTALL)
        if product_code_match:
            results['product_id'] = product_code_match.group(1).strip()

        revision_match = re.search(r'Revision: ([^\n]+)', output, re.DOTALL)
        if revision_match:
            results['esi_revision'] = revision_match.group(1).strip()

        ECAT_add = re.search(r'ECAT Address: ([^\n]+)', output, re.DOTALL)
        if ECAT_add:
            results['ethercat_address'] = ECAT_add.group(1).strip()

        version_match = re.search(r'Software version: ([^\n]+)', output)
        if version_match:
            results['firmware_version'] = version_match.group(1).strip()

        # Extract OTP programmed values
        pcb_pn_match = re.search(r'PCB_Part_Number:([^\s]+)', output)
        if pcb_pn_match:
            results['pcb_part_number'] = pcb_pn_match.group(1).strip()

        pcb_sn_match = re.search(r'PCB_Serial_Number:([^\s]+)', output)
        if pcb_sn_match:
            results['pcb_serial_number'] = pcb_sn_match.group(1).strip()

        assy_pn_match = re.search(r'Assembly_Part_Number:([^\s]+)', output)
        if assy_pn_match:
            results['assembly_part_number'] = assy_pn_match.group(1).strip()

        assy_sn_match = re.search(r'Assembly_Serial_Number:([^\s]+)', output)
        if assy_sn_match:
            results['assembly_serial_number'] = assy_sn_match.group(1).strip()

        return results

    def append_console_message(self, message, is_error=False):
        """Helper method to append colored messages to console"""
        if hasattr(self, 'console_output') and self.console_output is not None:
            if is_error:
                self.console_output.append(f'<span style="color:red; font-weight:bold;">{message}</span>')
            else:
                self.console_output.append(f'<span style="color:green;font-weight:bold;">{message}</span>')
            # Auto-scroll to bottom
            self.console_output.verticalScrollBar().setValue(
                self.console_output.verticalScrollBar().maximum()
            )

    def handle_ssh_error(self, error_msg):
        if hasattr(self, 'console_output') and self.console_output is not None:
            self.append_console_message(f"!!! SSH ERROR !!!\n{error_msg}\n", is_error=True)

        QMessageBox.critical(self, "SSH Error", error_msg)

    def VNA_cal_test(self):
        try:

            if hasattr(self, 'worker') and self.worker:
                self.worker.stop()

            self.vna_t = time.time()
            # Reset UI state
            self.VNAtest_console.clear()
            self.VNA_status_label_start.setText("Test: Running...")
            self.VNA_status_label_start.setStyleSheet("""
                QLabel {
                    background-color: #ffc107;
                    color: black;
                    padding: 2px 5px;
                    border-radius: 3px;
                    font-weight: bold;
                    font-size: 9pt;
                }
            """)
            self.start_vna_progress()
            self.VNA_start_button.setEnabled(False)

            self.append_vna_message("\n================== VNA CAL Test Started =======================\n")

            self.worker = Worker(
                self.ssh_handler,
                '/home/robot/Manufacturing_test/aipc_beta/vnacalibration.py',
                f'dimm 50000 70000',
                350
            )
            self.worker.output_ready.connect(self.handle_vna_output)
            self.worker.finished_signal.connect(self.on_vna_test_finished)
            self.worker.error_occurred.connect(self.handle_vna_error)

            # Start the thread
            self.worker.start()

        except Exception as e:
            self.append_vna_message(f"Failed to start test: {str(e)}", is_error=True)
            self.cleanup_resources()

    def handle_vna_output(self, line):
        try:
            self.append_vna_message(f"{line}\n")
            if "no ping" in line:
                self.append_vna_message(f"VNA not connected to the Network", is_error=True)
                self.worker.stop()
                self.VNA_start_button.setEnabled(True)

            if "Calibration PASS" in line:
                self.VNA_status_label_start.setText("Test: Passed")
                self.VNA_status_label_start.setStyleSheet("""
                          QLabel {
                              background-color: #28a745;
                              color: white;
                              padding: 2px 5px;
                              border-radius: 3px;
                              font-weight: bold;
                              font-size: 9pt;
                          }
                      """)
                self.append_vna_message("\n=== VNA Calibration PASSED ===")
                self.vna_timer.stop()
                self.worker.stop()
                self.VNA_start_button.setEnabled(True)
                # self.work_timeout = 30

            elif "Calibration FAIL" in line :
                self.VNA_status_label_start.setText("Test: Failed")
                self.VNA_status_label_start.setStyleSheet("""
                          QLabel {
                              background-color: #dc3545;
                              color: white;
                              padding: 2px 5px;
                              border-radius: 3px;
                              font-weight: bold;
                              font-size: 9pt;
                          }
                      """)
                self.append_vna_message("\n!!! VNA Calibration FAILED !!!", is_error=True)

                self.vna_timer.stop()
                self.worker.stop()
                self.VNA_start_button.setEnabled(True)
                # self.work_timeout = 30
            elif "ERROR: Connect ECal module" in line :
                self.VNA_status_label_start.setText("Test: Failed")
                self.VNA_status_label_start.setStyleSheet("""
                                          QLabel {
                                              background-color: #dc3545;
                                              color: white;
                                              padding: 2px 5px;
                                              border-radius: 3px;
                                              font-weight: bold;
                                              font-size: 9pt;
                                          }
                                      """)
                self.append_vna_message("\n!!! VNA Calibration FAILED : Please connect Ecal Module... !!!", is_error=True)
                self.vna_timer.stop()
                self.worker.stop()
                self.VNA_start_button.setEnabled(True)

        except Exception as e:
            # self.work_timeout = 30
            self.append_vna_message(f"Error processing output: {str(e)}", is_error=True)
            self.VNA_start_button.setEnabled(True)

    def on_vna_test_finished(self):
        self.vna_timer.stop()
        self.VNA_start_button.setEnabled(True)
        if hasattr(self, 'worker') and self.worker:
            self.worker.stop()

        # If test didn't explicitly pass or fail, mark it as incomplete
        if "Passed" not in self.VNA_status_label_start.text() and "Failed" not in self.VNA_status_label_start.text():
            self.VNA_status_label_start.setText("Test: Incomplete")
            self.VNA_status_label_start.setStyleSheet("""
                  QLabel {
                      background-color: #6c757d;
                      color: white;
                      padding: 2px 5px;
                      border-radius: 3px;
                      font-weight: bold;
                      font-size: 9pt;
                  }
              """)
            self.append_vna_message("\n!!! Test did not complete properly !!!", is_error=True)
        # self.work_timeout = 30

    def handle_vna_error(self, error_msg):
        self.vna_timer.stop()
        # self.work_timeout = 30
        self.cleanup_resources()
        self.append_vna_message(f"\n!!!  ERROR: {error_msg} !!!", is_error=True)
        self.VNA_start_button.setEnabled(True)

    def program_otp(self):
        if hasattr(self, 'console_output') and self.console_output is not None:
            self.console_output.clear()
        pcb_part_number = self.pcb_pn_input.text().strip()
        assembly_part_number = self.assembly_pn_input.text().strip()
        assembly_ser = self.assembly_sn_input.text().strip()
        assembly_rev = self.assembly_rev_input.text().strip()
        PCB_rev = self.pcb_rev_input.text().strip()
        PCB_ser = self.pcb_sn_input.text().strip()
        if len(pcb_part_number) == 0:
            self.append_console_message(f"ERROR: PCB part number not Entered\n", is_error=True)
            QMessageBox.critical(self, "Error", f"PCB part number not Entered")
            self.logger.error(f"PCB part number not Entered",
                              extra={'func_name': 'auto_load_connect'})
            return
        if len(PCB_rev) == 0:
            self.append_console_message(f"ERROR: PCB Revision not Entered\n", is_error=True)
            QMessageBox.critical(self, "Error", f"PCB Revision not Entered")
            self.logger.error(f"PCB Revision not Entered",
                              extra={'func_name': 'auto_load_connect'})
            return
        if len(assembly_ser) == 0:
            self.append_console_message(f"ERROR: Assembly Serial number not Entered\n", is_error=True)
            QMessageBox.critical(self, "Error", f"Assembly Serial number not Entered")
            # logger.error("Assembly Serial number not Entered")
            self.logger.error(f"Assembly Serial number not Entered",
                              extra={'func_name': 'auto_load_connect'})
            return
        if len(assembly_rev) == 0:
            self.append_console_message(f"ERROR: Assembly Revision not Entered\n", is_error=True)
            QMessageBox.critical(self, "Error", f"Assembly Revision not Entered")
            # logger.error("Assembly Revision not Entered")
            self.logger.error(f"Assembly Revision not Entered",
                              extra={'func_name': 'auto_load_connect'})
            return

        if len(PCB_ser) == 0:
            self.append_console_message(f"ERROR: PCB Serial not Entered\n", is_error=True)
            QMessageBox.critical(self, "Error", f"PCB Serial not Entered")
            self.logger.error(f"PCB Serial not Entered",
                              extra={'func_name': 'auto_load_connect'})
            return

        if len(assembly_part_number) == 0:
            self.append_console_message(f"ERROR :Assembly part number  not Entered \n", is_error=True)
            QMessageBox.critical(self, "Error", f"Assembly part number  not Entered")
            self.logger.error(f"Assembly part number  not Entered",
                              extra={'func_name': 'auto_load_connect'})
            return

        if not pcb_part_number or not self.validate_part_number(pcb_part_number, "PCB"):
            self.append_console_message(f"ERROR: Invalid PCB Part Number: {pcb_part_number}\n", is_error=True)
            self.logger.error(f"ERROR: Invalid PCB Part Number: {pcb_part_number}",
                              extra={'func_name': 'auto_load_connect'})
            return

        if not PCB_rev or not self.validate_revision_number(PCB_rev, "PCB"):
            self.append_console_message(f"ERROR: Invalid PCB Revision Number: {PCB_rev}\n", is_error=True)
            self.logger.error(f"ERROR: Invalid PCB Revision : {PCB_rev}",
                              extra={'func_name': 'auto_load_connect'})
            return

        if not assembly_part_number or not self.validate_part_number(assembly_part_number, "Assembly"):
            self.append_console_message(f"ERROR: Invalid Assembly Part Number: {assembly_part_number}\n",
                                        is_error=True)
            self.logger.error(f"ERROR: Invalid Assembly Part Number: {assembly_part_number}",
                              extra={'func_name': 'auto_load_connect'})
            return

        if not assembly_rev or not self.validate_revision_number(assembly_rev, "Assembly"):
            self.append_console_message(f"ERROR: Invalid Assembly Revision Number: {assembly_rev}\n", is_error=True)
            self.logger.error(f"ERROR: Invalid Assembly Revision : {assembly_rev}",
                              extra={'func_name': 'auto_load_connect'})
            return

        #self.append_console_message("Part numbers validated successfully. Connecting...\n\n")
        QApplication.processEvents()  # Update UI


        self.create_otp_file("C:\\tmp\\APPOTP", pcb_part_number, PCB_ser, assembly_part_number,assembly_ser)
        self.file_transer("C:\\tmp\\APPOTP","/home/robot/Manufacturing_test/aipc_beta/APPOTP")
        time.sleep(1)
        success, message = self.ssh_handler.Connect_RPI()
        if not success:
            self.handle_ssh_error(f"Connection failed: {message}")
            return

        self.execute_command("programotp", self.handle_otp_test_output, 0)

    def handle_otp_test_output(self,stdout, stderr):
        test_passed = "UPDATE_PASS" in stdout
        test_details = stdout.strip()

        if "Error in slave initialization" in test_details:
            QMessageBox.critical(
                self,
                "Critical Error",
                "Slave initialization failed! Please check the EtherCAT connection and restart the test."
            )
            self.test_status_label_start.setText("Failed")

        if test_passed:
            self.append_console_message("OTP Update PASS")
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(f"Insert the SD card and update the Firmware")
            msg.setWindowTitle("Firmware Update")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            retval = msg.exec_()

        else:
            self.append_console_message("OTP Update FAIL", is_error=True)



    def auto_load_connect(self):
        self.test_result = "PASS"
        if hasattr(self, 'console_output') and self.console_output is not None:
            self.console_output.clear()

        try:
            # self.config = self.load_config()
            
            if self.unit_test >= 1:
                excel_logger.reset_sheet("Unit Setup")
            self.unit_test += 1
            # Validate PCB Part Number
            # logger.info("Unit SETUP Test")
            self.auto_load_btn.setEnabled(False)
            self.append_console_message("==========TEST Started=============\n")
            Fixture = self.Fixture.text().strip()
            Testing_name = self.Test_name.text().strip()
            Ecal = self.Ecal_SN.text().strip()
            VNA = self.VNA_SN.text().strip()
            Vendor_name = self.Vendor_name.text().strip()
            pcb_part_number = self.pcb_pn_input.text().strip()
            assembly_part_number = self.assembly_pn_input.text().strip()
            assembly_ser = self.assembly_sn_input.text().strip()
            assembly_rev = self.assembly_rev_input.text().strip()
            PCB_rev = self.pcb_rev_input.text().strip()
            PCB_ser = self.pcb_sn_input.text().strip()
            self.test_result = 'PASS'
            self.PN = assembly_part_number
            self.SN = assembly_ser

            self.append_console_message("1. Validate Entry check\n")

            if len(pcb_part_number) == 0:
                self.append_console_message(f"ERROR: PCB part number not Entered\n", is_error=True)
                QMessageBox.critical(self, "Error", f"PCB part number not Entered")
                self.logger.error(f"PCB part number not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return
            if len(PCB_rev) == 0:
                self.append_console_message(f"ERROR: PCB Revision not Entered\n", is_error=True)
                QMessageBox.critical(self, "Error", f"PCB Revision not Entered")
                self.logger.error(f"PCB Revision not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return
            if len(assembly_ser) == 0:
                self.append_console_message(f"ERROR: Assembly Serial number not Entered\n", is_error=True)
                QMessageBox.critical(self, "Error", f"Assembly Serial number not Entered")
                # logger.error("Assembly Serial number not Entered")
                self.logger.error(f"Assembly Serial number not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return
            if len(assembly_rev) == 0:
                self.append_console_message(f"ERROR: Assembly Revision not Entered\n", is_error=True)
                QMessageBox.critical(self, "Error", f"Assembly Revision not Entered")
                # logger.error("Assembly Revision not Entered")
                self.logger.error(f"Assembly Revision not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if len(PCB_ser) == 0:
                self.append_console_message(f"ERROR: PCB Serial not Entered\n", is_error=True)
                QMessageBox.critical(self, "Error", f"PCB Serial not Entered")
                self.logger.error(f"PCB Serial not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if len(assembly_part_number) == 0:
                self.append_console_message(f"ERROR :Assembly part number  not Entered \n", is_error=True)
                QMessageBox.critical(self, "Error", f"Assembly part number  not Entered")
                self.logger.error(f"Assembly part number  not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if len(Vendor_name) == 0:
                self.append_console_message(f"ERROR :Vendor Name not Entered \n", is_error=True)
                QMessageBox.critical(self, "Error", f"Vendor Name not Entered")
                self.logger.error(f"Vendor Name  not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if len(Fixture) == 0:
                self.append_console_message(f"ERROR :Fixture Number not Entered \n", is_error=True)
                QMessageBox.critical(self, "Error", f"Fixture Number not Entered")
                self.logger.error(f"Fixture Number not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if len(Testing_name) == 0:
                self.append_console_message(f"ERROR: Test Name not Entered\n", is_error=True)
                QMessageBox.critical(self, "Error", f"Test Operator Name not Entered")
                self.logger.error(f"Test Operator Name not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return
            if len(Ecal) == 0:
                self.append_console_message(f"ERROR : ECAL SN not Entered \n", is_error=True)
                QMessageBox.critical(self, "Error", f"ECAL SN not Entered")
                self.logger.error(f"ECAL SN not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return
            if len(VNA) == 0:
                self.append_console_message(f"ERROR: VNA SN not Entered\n", is_error=True)
                QMessageBox.critical(self, "Error", f"VNA SN not Entered")
                self.logger.error(f"VNA SN not Entered",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if not pcb_part_number or not self.validate_part_number(pcb_part_number, "PCB"):
                self.append_console_message(f"ERROR: Invalid PCB Part Number: {pcb_part_number}\n", is_error=True)
                self.logger.error(f"ERROR: Invalid PCB Part Number: {pcb_part_number}",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if not PCB_rev or not self.validate_revision_number(PCB_rev, "PCB"):
                self.append_console_message(f"ERROR: Invalid PCB Revision Number: {PCB_rev}\n", is_error=True)
                self.logger.error(f"ERROR: Invalid PCB Revision : {PCB_rev}",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if not assembly_part_number or not self.validate_part_number(assembly_part_number, "Assembly"):
                self.append_console_message(f"ERROR: Invalid Assembly Part Number: {assembly_part_number}\n",
                                            is_error=True)
                self.logger.error(f"ERROR: Invalid Assembly Part Number: {assembly_part_number}",
                                  extra={'func_name': 'auto_load_connect'})
                return

            if not assembly_rev or not self.validate_revision_number(assembly_rev, "Assembly"):
                self.append_console_message(f"ERROR: Invalid Assembly Revision Number: {assembly_rev}\n", is_error=True)
                self.logger.error(f"ERROR: Invalid Assembly Revision : {assembly_rev}",
                                  extra={'func_name': 'auto_load_connect'})
                return



            self.append_console_message("Part numbers validated successfully. Connecting...\n\n")
            QApplication.processEvents()  # Update UI

            assembly_suffix = self._extract_assembly_suffix(assembly_part_number)
            self.assembly_suffix = assembly_suffix
            if assembly_suffix:
                self.append_console_message(f"Using configuration for assembly suffix: {assembly_suffix}\n")
                if self.config_transfer(assembly_suffix) == False:
                    return
                self.config = self.load_config(assembly_suffix)
                

                

            else:
                self.append_console_message(
                    f"WARNING: Assembly part number suffix not recognized.Please check the part no and  update the config file\n", is_error=True)
                return
                #self.config_transfer()  # Use default config


            # Connect to SSH
            success, message = self.ssh_handler.Connect_RPI()
            if not success:
                self.handle_ssh_error(f"Connection failed: {message}")
                return

            # Execute commands sequentially
            commands = [
                ("soemcompile", self.handle_soemcompile_output),
                ("firmwarecheck", self.handle_firmare_check_output),
                ("otpcheck", self.handle_otpcheck_output),
                ("slaveinfo", self.handle_slaveinfo_output)
            ]
            val = 2
            for cmd, handler in commands:
                if not self.execute_command(cmd, handler, val):
                    if cmd != 'firmwarecheck':
                        break  # Stop if any command fails
                    if self.Firmware_check == False:
                        break

                val = val + 1
            self.append_console_message(
                "======================== Unit Setup completed====================================\n")
            # Log metadata along with test data
            self.excel_logger.log_summary(
                metadata={
                    'eid': f'{self.assembly_pn_input.text().strip()}_{self.assembly_rev_input.text().strip()}_{self.assembly_sn_input.text().strip()[-3:]}',
                    'serial_number': self.assembly_sn_input.text().strip(),
                    'model_number': 'AIPC I-BOX',
                    'version': 'V1.0',
                    'tester_name': self.Test_name.text().strip(),
                    'comment': 'Control Limit',
                    'start_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'end_time': '',  # Will be filled when test completes
                    'overall_result': '',  # Will be updated to PASS/FAIL
                    'test_fixture_sn': self.Fixture.text().strip(),
                    'vna_sn': self.VNA_SN.text().strip(),
                    'ecal_sn':self.Ecal_SN.text().strip()
                }
            )

            unit_data = {
                'Vendor_name': self.Vendor_name.text().strip(),
                'Fixture_number': self.Fixture.text().strip(),
                'test_operator_name': self.Test_name.text().strip(),
                'test_date': self.Test_Date.date().toString("yyyy-MM-dd"),
                'vna_calibration_date': self.VNA_calibration.date().toString("yyyy-MM-dd"),
                'vna_sn': self.VNA_SN.text().strip(),
                'ecal_sn': self.Ecal_SN.text().strip(),
                'pcb_part_number': self.pcb_pn_input.text().strip(),
                'pcb_revision': self.pcb_rev_input.text().strip(),
                'pcb_serial_number': self.pcb_sn_input.text().strip(),
                'assembly_part_number': self.assembly_pn_input.text().strip(),
                'assembly_revision': self.assembly_rev_input.text().strip(),
                'assembly_serial_number': self.assembly_sn_input.text().strip(),
                'product_id': self.product_id.text().strip(),
                'esi_revision': self.esi_revision.text().strip(),
                'configuration_id': self.configuration_id.text().strip(),
                'ethercat_address': self.ethercat_address.text().strip(),
                'firmware_version': self.firmware_version.text().strip()
            }
            self.excel_logger.log_unit_setup(unit_data)
            self.excel_logger.update_overall_result(self.test_result, PN=self.PN, SN=self.SN)

        except Exception as e:
            # self.logger.error(f"Error in auto_load_connect: {str(e)}",exc_info=True,extra={'func_name': 'auto_load_connect'} )
            self.auto_load_btn.setEnabled(True)
            QMessageBox.critical(self, "Error", f"Auto load failed: {str(e)}")
        finally:
            self.auto_load_btn.setEnabled(True)
            self.ssh_handler.SSH_disconnect()

    def dimm_cal_test(self):
        try:

            if hasattr(self, 'worker') and self.worker:
                self.worker.stop()

            # Reset UI state
            self.dimmtest_console.clear()
            self.DIMM_status_label_start.setText("Test: Running...")
            self.DIMM_status_label_start.setStyleSheet("""
                QLabel {
                    background-color: #ffc107;
                    color: black;
                    padding: 2px 5px;
                    border-radius: 3px;
                    font-weight: bold;
                    font-size: 9pt;
                }
            """)
            self.start_dimm_progress()
            self.dimm_start_button.setEnabled(False)

            self.append_dimm_message("\n================== Dimm Test Started =======================\n")

            self.worker = Worker(
                self.ssh_handler,
                '/home/robot/Manufacturing_test/aipc_beta/dimmcalibration.py',
                'dimm'
            )
            self.worker.output_ready.connect(self.handle_dimm_output)
            self.worker.finished_signal.connect(self.on_dimm_test_finished)
            self.worker.error_occurred.connect(self.handle_dimm_error)

            # Start the thread
            self.worker.start()

        except Exception as e:
            self.append_dimm_message(f"Failed to start test: {str(e)}", is_error=True)
            self.cleanup_resources()

    def handle_dimm_output(self, line):
        try:
            if "no ping" in line:
                self._log_Impedance_message(f"DIMM not connected to the Network", is_error=True)
                self.dimm_timer.stop()
                self.worker.stop()
                self.dimm_start_button.setEnabled(True)
            if "Calibration Pass" in line:
                self.DIMM_status_label_start.setText("Test: Passed")
                self.DIMM_status_label_start.setStyleSheet("""
                        QLabel {
                            background-color: #28a745;
                            color: white;
                            padding: 2px 5px;
                            border-radius: 3px;
                            font-weight: bold;
                            font-size: 9pt;
                        }
                    """)
                self.append_dimm_message("\n=== DIMM Calibration PASSED ===")

                # Log to Excel
                """ 
                unit_identifier = f"{self.assembly_pn_input.text().strip()} ({self.assembly_sn_input.text().strip()})"
                self.excel_logger.log_self_test(
                    unit_identifier=unit_identifier,
                    test_passed=True,
                    test_details="DIMM Calibration",
                    notes=line
                )"""
                self.dimm_timer.stop()
                self.worker.stop()
                self.dimm_start_button.setEnabled(True)

            elif "Calibration Fail" in line:
                self.DIMM_status_label_start.setText("Test: Failed")
                self.DIMM_status_label_start.setStyleSheet("""
                        QLabel {
                            background-color: #dc3545;
                            color: white;
                            padding: 2px 5px;
                            border-radius: 3px;
                            font-weight: bold;
                            font-size: 9pt;
                        }
                    """)
                self.append_dimm_message("\n!!! DIMM Calibration FAILED !!!", is_error=True)

                # Log to Excel
                """
                unit_identifier = f"{self.assembly_pn_input.text().strip()} ({self.assembly_sn_input.text().strip()})"
                self.excel_logger.log_self_test(
                    unit_identifier=unit_identifier,
                    test_passed=False,
                    test_details="DIMM Calibration",
                    notes=line
                )"""
                self.dimm_timer.stop()
                self.worker.stop()
                self.dimm_start_button.setEnabled(True)


        except Exception as e:
            self.append_dimm_message(f"Error processing output: {str(e)}", is_error=True)
            self.dimm_start_button.setEnabled(True)


    def on_dimm_test_finished(self):
        self.dimm_timer.stop()
        self.dimm_start_button.setEnabled(True)
        if hasattr(self, 'worker') and self.worker:
            self.worker.stop()

        # If test didn't explicitly pass or fail, mark it as incomplete
        if "Passed" not in self.DIMM_status_label_start.text() and "Failed" not in self.DIMM_status_label_start.text():
            self.DIMM_status_label_start.setText("Test: Incomplete")
            self.DIMM_status_label_start.setStyleSheet("""
                QLabel {
                    background-color: #6c757d;
                    color: white;
                    padding: 2px 5px;
                    border-radius: 3px;
                    font-weight: bold;
                    font-size: 9pt;
                }
            """)
            self.append_dimm_message("\n!!! Test did not complete properly !!!", is_error=True)


    def handle_dimm_error(self, error_msg):
        self.dimm_timer.stop()
        self.cleanup_resources()
        self.append_dimm_message(f"\n!!! ERROR: {error_msg} !!!", is_error=True)
        self.DIMM_status_label_start.setText("Test: Error")
        self.DIMM_status_label_start.setStyleSheet("""
            QLabel {
                background-color: #dc3545;
                color: white;
                padding: 2px 5px;
                border-radius: 3px;
                font-weight: bold;
                font-size: 9pt;
            }
        """)
        self.dimm_start_button.setEnabled(True)

        # Log to Excel
        unit_identifier = f"{self.assembly_pn_input.text().strip()} ({self.assembly_sn_input.text().strip()})"
        self.excel_logger.log_self_test(
            unit_identifier=unit_identifier,
            test_passed=False,
            test_details="DIMM Calibration Error",
            notes=error_msg
        )


    def append_dimm_message(self, message, is_error=False):
        """Helper method to append colored messages to DIMM console"""
        if hasattr(self, 'dimmtest_console') and self.dimmtest_console is not None:
            if is_error:
                self.dimmtest_console.append(f'<span style="color:red; font-weight:bold;">{message}</span>')
            else:
                self.dimmtest_console.append(f'<span style="color:green; font-weight:bold;">{message}</span>')

            # Auto-scroll to bottom
            self.dimmtest_console.verticalScrollBar().setValue(
                self.dimmtest_console.verticalScrollBar().maximum()
            )

    def append_vna_message(self, message, is_error=False):
        """Helper method to append colored messages to DIMM console"""
        if hasattr(self, 'VNAtest_console') and self.VNAtest_console is not None:
            if is_error:
                self.VNAtest_console.append(f'<span style="color:red; font-weight:bold;">{message}</span>')
            else:
                self.VNAtest_console.append(f'<span style="color:green; font-weight:bold;">{message}</span>')

            # Auto-scroll to bottom
            self.VNAtest_console.verticalScrollBar().setValue(
                self.VNAtest_console.verticalScrollBar().maximum()
            )

    def append_BNC_message(self, message, is_error=False):
        """Helper method to append colored messages to DIMM console"""
        if hasattr(self, 'BNCtest_console') and self.BNCtest_console is not None:
            if is_error:
                self.BNCtest_console.append(f'<span style="color:red; font-weight:bold;">{message}</span>')
            else:
                self.BNCtest_console.append(f'<span style="color:green; font-weight:bold;">{message}</span>')

            # Auto-scroll to bottom
            self.BNCtest_console.verticalScrollBar().setValue(
                self.BNCtest_console.verticalScrollBar().maximum()
            )

    def BNC_test(self):
        try:
            if self.bnc_t >= 1:
                excel_logger.reset_sheet("BNC Port Verification")
            self.bnc_t += 1
            self.overall_result = 'PASS'
            # self.start_time1 = time.time()

            if hasattr(self, 'worker') and self.worker:
                self.worker.stop()

            # Reset UI state
            self.BNCtest_console.clear()
            self.BNC_status_label_start.setText("Test: Running...")
            self.BNC_status_label_start.setStyleSheet("""
            QLabel {
                    background-color: #ffc107;
                    color: black;
                    padding: 2px 5px;
                    border-radius: 3px;
                    font-weight: bold;
                    font-size: 9pt;
            }"""
                                                      )
            self.BNC_start_button.setEnabled(False)

            self.append_BNC_message("\n================== BNC Test Started =======================\n")

            # Show first prompt for Zone 1
            self.show_zone_prompt(2)

        except Exception as e:
            self.append_BNC_message(f"Failed to start test: {str(e)}", is_error=True)
            self.cleanup_resources()

    def show_zone_prompt(self, zone_number):

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(f"Please connect Zone {zone_number} and click OK to continue")
        msg.setWindowTitle("Zone Connection")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        command = f'{zone_number} dimm'

        retval = msg.exec_()

        if retval == QMessageBox.Ok:
            self.start_time1 = time.time()
            self.append_BNC_message(f"\nTesting Zone {zone_number}...\n")

            # Start the worker with the main script
            self.worker = Worker(
                self.ssh_handler,
                '/home/robot/Manufacturing_test/aipc_beta/BNC.py',  # Your original script
                command
            )
            # Connect the output handler
            self.worker.output_ready.connect(self.handle_BNC_output)
            self.worker.error_occurred.connect(self.handle_BNC_error)
            self.worker.start()
        else:
            self.append_BNC_message("Test cancelled by user", is_error=True)
            self.BNC_status_label_start.setText("Cancelled")
            self.BNC_start_button.setEnabled(True)
            self.BNC_status_label_start.setStyleSheet("background-color: #dc3545; color: white;")

    def handle_BNC_output(self, line):
        # Your original output handling - completely unchanged
        if "Error in slave initialization" in line:
            QMessageBox.critical(
                self,
                "Critical Error",
                "Slave initialization failed! Please check the EtherCAT connection and restart the test."
            )
            self.worker.stop()
        """
        if "Zone1-Inner" in line:
            Val = line.split(",")
            Testname = Val[0]
            Value = Val[1]
            Result = Val[2]

            if Result.upper() == "PASS":
                self.append_BNC_message(f"\nBNC Test Zone1-Inner PASS\n")
                val = True

            else:
                self.append_BNC_message(f"\nBNC Test Zone1-Inner Fail\n", is_error=True)
                val = False

                self.overall_result = 'FAIL'

            self.excel_logger.log_BNC_measurement(
                test_zone=Testname,
                test_details=Value,
                test_passed=val
            )
            self.step_no = self.step_no + 1
            # Log metadata along with test data
            self.excel_logger.log_summary(
                step_data={
                    'step': str(self.step_no),
                    'unit': 'dB',
                    'low_limit': '-1',
                    'measure': Value,
                    'high_limit': '0',
                    'teststep': 'Verify BNC port',
                    'testpoints': 'Zone1',
                    'status': "PASS" if val else "FAIL"
                }
            )

            # After Zone1 is done, prompt for Zone2

            self.worker.stop()
            time.sleep(2)
            self.show_zone_prompt(2)
            # self.handle_test_failure()
        """

        if "Zone2-Mid_Inner" in line:
            Val = line.split(",")
            if len(Val) < 3:
                self.append_BNC_message(f"Invalid data format for Zone2-Mid_Inner: {line}", is_error=True)
                return
            Testname = Val[0]
            Value = Val[1]
            Result = Val[2]
            print(Val)

            if Result.upper() == "PASS":
                self.append_BNC_message("BNC Test Zone2-Mid_Inner PASS")
                val = True

            else:
                self.append_BNC_message(f"BNC Test Zone2-Mid_Inner Fail", is_error=True)
                val = False
                self.overall_result = 'FAIL'


            self.excel_logger.log_BNC_measurement(
                test_zone=Testname,
                test_details=Value,
                test_passed=val
            )
            self.step_no = self.step_no + 1
            self.excel_logger.log_summary(
                step_data={
                    'step': str(self.step_no),
                    'unit': 'dB',
                    'low_limit': '-1',
                    'measure': Value,
                    'high_limit': '0',
                    'teststep': 'Verify BNC port',
                    'testpoints': 'Zone2',
                    'status': "PASS" if val else "FAIL"
                }
            )

            # After Zone2 is done, prompt for Zone3
            self.worker.stop()
            self.show_zone_prompt(3)
            # self.handle_test_failure()

        elif "Zone3-Mid_Edge" in line:
            Val = line.split(",")
            if len(Val) < 3:
                self.append_BNC_message(f"Invalid data format for Zone3-Mid_Edge: {line}", is_error=True)
                return
            Testname = Val[0]
            Value = Val[1]
            Result = Val[2]

            if Result.upper() == "PASS":
                self.append_BNC_message(f"\nBNC Test Zone3-Mid_Edge PASS\n")
                val = True

            else:
                self.append_BNC_message(f"\nBNC Test Zone3-Mid_Edge Fail\n", is_error=True)
                val = False
                self.overall_result = 'FAIL'


            self.excel_logger.log_BNC_measurement(
                test_zone=Testname,
                test_details=Value,
                test_passed=val
            )
            self.step_no= self.step_no+ 1
            self.excel_logger.log_summary(
                step_data={
                    'step': str(self.step_no),
                    'unit': 'dB',
                    'low_limit': '-1',
                    'measure': Value,
                    'high_limit': '0',
                    'teststep': 'Verify BNC port',
                    'testpoints': 'Zone3',
                    'status': "PASS" if val else "FAIL"
                }
            )

            self.worker.stop()
            self.show_zone_prompt(4)

            # self.handle_test_failure()

        elif "Zone4-Edge" in line:
            Val = line.split(",")
            if len(Val) < 3:
                self.append_BNC_message(f"Invalid data format for Zone4-Edge: {line}", is_error=True)
                return
            Testname = Val[0]
            Value = Val[1]
            Result = Val[2]

            if Result.upper() == "PASS":
                self.append_BNC_message(f"\nBNC Test Zone4-Edge PASS\n")
                val = True

            else:
                self.append_BNC_message(f"\nBNC Test Zone4-Edge Fail\n", is_error=True)
                val = False
                self.overall_result = 'FAIL'


            self.excel_logger.log_BNC_measurement(
                test_zone=Testname,
                test_details=Value,
                test_passed=val
            )
            self.step_no = self.step_no + 1
            self.excel_logger.log_summary(
                step_data={
                    'step': str(self.step_no),
                    'unit': 'dB',
                    'low_limit': '-1',
                    'measure': Value,
                    'high_limit': '0',
                    'teststep': 'Verify BNC port',
                    'testpoints': 'Zone4',
                    'status': "PASS" if val else "FAIL"
                }
            )

            self.worker.stop()
            self.show_zone_prompt(5)

            # self.handle_test_failure()

        elif "Zone5-Outer" in line:
            Val = line.split(",")
            if len(Val) < 3:
                self.append_BNC_message(f"Invalid data format for Zone5-Outer: {line}", is_error=True)
                return
            Testname = Val[0]
            Value = Val[1]
            Result = Val[2]

            if Result.upper() == "PASS":
                self.append_BNC_message(f"\nBNC Test Zone5-Outer PASS\n")
                val = True

            else:
                self.append_BNC_message(f"\nBNC Test Zone5-Outer Fail\n", is_error=True)
                val = False

                self.overall_result = 'FAIL'
                self.over_all_result = 'FAIL'


            self.excel_logger.log_BNC_measurement(
                test_zone=Testname,
                test_details=Value,
                test_passed=val
            )

            self.step_no = self.step_no + 1

            self.excel_logger.log_summary(
                step_data={
                    'step': str(self.step_no),
                    'unit': 'dB',
                    'low_limit': '-1',
                    'measure': Value,
                    'high_limit': '0',
                    'teststep': 'Verify BNC port',
                    'testpoints': 'Zone5',
                    'status': "PASS" if val else "FAIL"
                }
            )

            self.excel_logger.log_summary(
                metadata={
                    'end_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'overall_result': self.over_all_result
                }
            )

            self.excel_logger.update_overall_result(self.over_all_result)



            self.append_BNC_message(f"BNC Test completed successfully")
            self.worker.stop()
            self.BNC_status_label_start.setText("Completed")
            self.BNC_start_button.setEnabled(True)
            self.BNC_status_label_start.setStyleSheet("background-color: #28a745; color: white;")

        if time.time() - self.start_time1 > 90:
            self.append_BNC_message(
                f"===No Data from Raspberry pi for more than 90 sec please check the Raspberry pi =====", is_error=True)
            self.worker.stop()

    def handle_BNC_error(self, error_msg):
        self.append_BNC_message(f"ERROR: {error_msg}", is_error=True)
        self.BNC_start_button.setEnabled(True)
        self.cleanup_resources()
        self.BNC_status_label_start.setText("Failed")
        self.BNC_status_label_start.setStyleSheet("background-color: #dc3545; color: white;")


    @log_function
    def execute_command(self, command, output_handler, val):
        """Execute a single command and handle its output"""

        if command != "selftest":
            self.append_console_message(f"\n{str(val)}. {command}\n")
        QApplication.processEvents()  # Update UI
        self.logger.info(f"Executing command: {command}",
                         extra={'func_name': command})
        try:
            stdout, stderr = self.ssh_handler.SSH_com(command)
            self.logger.info(f"Command output received",
                             extra={'func_name': command})

            if stdout:
                self.logger.debug(f"stdout:\n{stdout}",
                                  extra={'func_name': command})
            if stderr:
                self.logger.error(f"stderr:\n{stderr}",
                                  extra={'func_name': command})
            # Process the output
            if stderr and command != 'soemcompile':
                self.append_console_message(f"!!! ERROR !!!\n{stderr}\n", is_error=True)

            return output_handler(stdout, stderr)

        except Exception as e:
            self.logger.error(f"Command failed: {str(e)}",
                              exc_info=True,
                              extra={'func_name': command})
            self.handle_ssh_error(f"Command '{command}' failed: {str(e)}")
            return False

    def handle_otpcheck_output(self, stdout, stderr):
        self.handling_flag = 1

        if "No such file or directory" in stderr:
            error_msg = "ERROR: No OTP file found on device\n"
            self.console_output.append("!!! CRITICAL ERROR !!!\n")
            self.console_output.append(error_msg)
            QMessageBox.critical(self, "OTP Error",
                                 "No OTP file found on device!\n\n"
                                 "Please check if the device has proper OTP configuration.\n"
                                 "Expected file: /dev/otp/APPOTP")
            return False

        parsed_data = self.parse_ssh_output(stdout)

        # Compare OTP programmed values with GUI entered values
        gui_pcb_pn = self.pcb_pn_input.text().strip()
        gui_pcb_sn = self.pcb_sn_input.text().strip()
        gui_assy_pn = self.assembly_pn_input.text().strip()
        gui_assy_sn = self.assembly_sn_input.text().strip()

        otp_pcb_pn = parsed_data.get('pcb_part_number', '')
        otp_pcb_pn = otp_pcb_pn.split('_')[0] if otp_pcb_pn else ''
        otp_pcb_sn = parsed_data.get('pcb_serial_number', '')
        otp_assy_pn = parsed_data.get('assembly_part_number', '')
        otp_assy_sn = parsed_data.get('assembly_serial_number', '')

        match_failures = []

        if gui_pcb_pn and otp_pcb_pn and gui_pcb_pn != otp_pcb_pn:
            match_failures.append(f"PCB Part Number mismatch (GUI: {gui_pcb_pn} vs OTP: {otp_pcb_pn})")

        if gui_pcb_sn and otp_pcb_sn and gui_pcb_sn != otp_pcb_sn:
            match_failures.append(f"PCB Serial Number mismatch (GUI: {gui_pcb_sn} vs OTP: {otp_pcb_sn})")

        if gui_assy_pn and otp_assy_pn and gui_assy_pn != otp_assy_pn:
            match_failures.append(f"Assembly Part Number mismatch (GUI: {gui_assy_pn} vs OTP: {otp_assy_pn})")

        if gui_assy_sn and otp_assy_sn and gui_assy_sn != otp_assy_sn:
            match_failures.append(f"Assembly Serial Number mismatch (GUI: {gui_assy_sn} vs OTP: {otp_assy_sn})")

        if match_failures:
            error_msg = "OTP Programming Verification Failed:\n" + "\n".join(match_failures)
            self.append_console_message("!!! OTP PROGRAMMING ERROR !!!\n", is_error=True)
            self.append_console_message(error_msg + "\n", is_error=True)
            self.test_result = 'FAIL'
            return False
        else:
            success_msg = "OTP Programming Verified Successfully!\n"
            success_msg += f"PCB Part Number: {otp_pcb_pn}\n"
            success_msg += f"PCB Serial Number: {otp_pcb_sn}\n"
            success_msg += f"Assembly Part Number: {otp_assy_pn}\n"
            success_msg += f"Assembly Serial Number: {otp_assy_sn}\n"

            # self.append_console_message("=== OTP VERIFICATION SUCCESS ===\n")
            self.append_console_message(success_msg)
            # QMessageBox.information(self, "OTP Verified", success_msg)
            return True

    def handle_firmare_check_output(self, stdout, sterr):
        parsed_data = self.parse_ssh_output(stdout)
        Actual_version = parsed_data['firmware_version']
        Expected_version = self.config["expected_firmware_version"]
        if Actual_version == Expected_version:
            self.Firmware_check = True
            self.append_console_message(
                f"✔ Correct version: {Expected_version}")
        else:
            error_msg = "Firmware Mismatch Test cannot be proceed further please update the latest Firmware"
            self.Firmware_check = False
            self.append_console_message(
                f"✖ Incorrect version (expected: {Expected_version} Actual: {Actual_version})",
                is_error=True
            )
            QMessageBox.critical(self, "Firmware Version Mismatch", error_msg)
            self.test_result = 'FAIL'

    def handle_self_test_output(self, stdout, sterr):
        test_passed = "Self Test PASS" in stdout
        test_details = stdout.strip()

        if "Error in slave initialization" in  test_details:
            QMessageBox.critical(
                self,
                "Critical Error",
                "Slave initialization failed! Please check the EtherCAT connection and restart the test."
            )
            self.test_status_label_start.setText("Failed")


        if test_passed:
            self.append_self_message("SELF TEST PASS")
            self.test_status_label_start.setText("Passed")
            self.test_status_label_start.setStyleSheet("background-color: #28a745; color: white;")
        else:
            self.append_self_message("SELF TEST FAIL", is_error=True)
            self.over_all_result = "FAIL"
            self.test_status_label_start.setText("Failed")
            self.test_status_label_start.setStyleSheet("background-color: #dc3545; color: white;")

        # Log to Excel
        #unit_identifier = f"{self.assembly_pn_input.text().strip()} ({self.assembly_sn_input.text().strip()})"
        self.excel_logger.log_self_test(
            unit_identifier="self_test",
            test_passed=test_passed,
            test_details=test_details,
            notes="Self Test completed"
        )

        self.excel_logger.log_summary(
            metadata={
                'end_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'overall_result': self.over_all_result
            }
        )

        return test_passed

    def check_output_for_strings(self, stdout):
        """Check the output for specific strings and return findings"""
        results = {
            'compilation_success': False,
            'pdo_map_success': False,
        }

        # Check for successful compilation
        if "Ethercat compiled Sucessfully" in stdout:
            results['compilation_success'] = True

        # Check for PDO map completion
        if "pdo map successfully reached end" in stdout:
            results['pdo_map_success'] = True

        return results

    # Example usage in your handle_soemcompile_output method:
    def handle_soemcompile_output(self, stdout, stderr):
        self.handling_flag = 2

        # Analyze the output
        analysis = self.check_output_for_strings(stdout)

        # Check for specific conditions
        if not analysis['compilation_success']:
            error_msg = "ERROR: Ethercat compilation failed\n"
            self.append_console_message("!!! CRITICAL ERROR !!!\n", is_error=True)
            self.append_console_message(error_msg, is_error=True)
            QMessageBox.critical(self, "Compilation Error", error_msg)
            self.test_result = 'FAIL'
            return False

        if not analysis['pdo_map_success']:
            error_msg = "ERROR: PDO Map generation failed\n"
            self.append_console_message("!!! CRITICAL ERROR !!!\n", is_error=True)
            self.append_console_message(error_msg, is_error=True)
            QMessageBox.critical(self, "PDO Map Error", error_msg)
            self.test_result = 'FAIL'
            return False

        if "Ethercat compiled Sucessfully" in stdout and "pdo map successfully reached end" in stdout:
            success_msg = "Ethercat compiled Successfully with PDO mapping"
            self.append_console_message(success_msg)
            self.check = True
            return True

        return False

    def handle_slaveinfo_output(self, stdout, stderr):
        Counter = 0
        if not self.check:
            return False

        self.handling_flag = 3
        parsed_data = self.parse_ssh_output(stdout)

        if len(parsed_data['product_id']) == 0 or parsed_data['product_id'] is None:
            self.append_console_message(" Error : Product ID is None", is_error=True)
            Counter = 1
        if len(parsed_data['esi_revision']) == 0 or parsed_data['esi_revision'] is None:
            self.append_console_message(" Error : ESI Revision is None", is_error=True)
            Counter = 1
        if len(parsed_data['ethercat_address']) == 0 or parsed_data['ethercat_address'] is None:
            self.append_console_message(" Error : Ethercat Address is None", is_error=True)
            Counter = 1

        if int(parsed_data['ethercat_address'],16) !=  int('0x444',16) :
            Address = parsed_data['ethercat_address']
            self.append_console_message(f" Error : Ethercat Address Failed Actual_Value = {Address} Expected_Value = 0X444  ", is_error=True)
            self.append_console_message(f"Set hex switch to 0x444 and Power Cycle and restart the test", is_error=True)
            Counter = 1

        self.product_id.setText(parsed_data['product_id'])
        self.esi_revision.setText(parsed_data['esi_revision'])
        self.firmware_version.setText(parsed_data['firmware_version'])
        self.ethercat_address.setText(parsed_data['ethercat_address'])

        if Counter == 0:
            self.append_console_message("Slave Information Test Passed")
        else:
            self.append_console_message("Slave Information Test Fail", is_error=True)
            self.test_result = 'FAIL'

        return True

    def create_form_row(self, label, widget, layout, readonly=False):
        """Create a responsive form row"""
        row = QWidget()
        row_layout = QVBoxLayout(row)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(2)

        lbl = QLabel(label)
        lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        row_layout.addWidget(lbl)

        if isinstance(widget, (QLineEdit, QDateEdit)):
            widget.setMinimumHeight(30)  # Reduced from 38 for better scaling
            if readonly:
                widget.setReadOnly(True)
                if isinstance(widget, QLineEdit):
                    widget.setStyleSheet("""
                              background-color: #e0e0e0;
                              color: #495057;
                              border: 1px solid #000000;
                              font-weight: bold;
                          """)

        widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        row_layout.addWidget(widget)
        layout.addWidget(row)
        return widget

    def _create_Impedance_zone_panel(self, zone_name):
        """Create a responsive impedance zone panel"""
        panel = QGroupBox(zone_name)
        panel.setFont(QFont('Arial', 10, QFont.Bold))

        panel.setStyleSheet("""
              QGroupBox {
                  border: 2px solid #aaa;
                  border-radius: 5px;
                  margin-top: 10px;
                  padding-top: 18px;
                  background: #f8f8f8;
              }
              QGroupBox::title {
                  subcontrol-origin: margin;
                  left: 10px;
                  padding: 0 5px;
              }
          """)

        panel_layout = QVBoxLayout(panel)
        panel_layout.setSpacing(10)
        panel_layout.setContentsMargins(8, 15, 8, 8)

        # Create measurement table with scroll area
        table_scroll = QScrollArea()
        table_scroll.setWidgetResizable(True)
        table_scroll.setMinimumHeight(200)

        measurement_table = QTableWidget(1 if zone_name == "Zone1-Inner" else len(self._relay_values_imp), 5)
        measurement_table.setHorizontalHeaderLabels(["Setpt", "R", "I", "Z", "P/F"])
        measurement_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        measurement_table.verticalHeader().setVisible(False)
        measurement_table.setFont(QFont('Arial', 9))

        # Style the table
        measurement_table.setStyleSheet("""
              QTableWidget {
                  background-color: white;
                  gridline-color: #e0e0e0;
              }
              QHeaderView::section {
                  background-color: #e0e0e0;
                  padding: 4px;
                  border: 1px solid #e0e0e0;
                  font-weight: bold;
              }
              QTableWidget::item {
                  background-color: #f5f5f5;
              }
          """)

        # Populate setpoint values
        if zone_name == "Zone1-Inner":
            measurement_table.setItem(0, 0, QTableWidgetItem(str(self._relay_values_imp[0])))
        else:
            for row, relay in enumerate(self._relay_values_imp):
                measurement_table.setItem(row, 0, QTableWidgetItem(str(relay)))

        table_scroll.setWidget(measurement_table)
        panel_layout.addWidget(table_scroll)

        # Add test button
        test_button = QPushButton(f"Test {zone_name}")
        test_button.setFont(QFont('Arial', 9))
        test_button.setFixedHeight(30)
        test_button.setStyleSheet("""
              QPushButton {
                  min-width: 80px;
                  padding: 4px;
                  background: #e0e0e0;
                  border: 1px solid #aaa;
              }
              QPushButton:hover {
                  background: #d0d0d0;
              }
          """)
        test_button.clicked.connect(lambda _, z=zone_name: self._start_impedance_zone_measurement(z))
        panel_layout.addWidget(test_button, alignment=Qt.AlignCenter)

        self._measurement_tables_imp[zone_name] = measurement_table

        return panel

    def _create_resistance_zone_panel(self, zone_name):
        """Create a responsive resistance zone panel"""
        panel = QGroupBox(zone_name)
        panel.setFont(QFont('Arial', 10, QFont.Bold))

        panel.setStyleSheet("""
               QGroupBox {
                   border: 2px solid #aaa;
                   border-radius: 5px;
                   margin-top: 10px;
                   padding-top: 18px;
                   background: #f8f8f8;
               }
               QGroupBox::title {
                   subcontrol-origin: margin;
                   left: 10px;
                   padding: 0 5px;
               }
           """)

        panel_layout = QVBoxLayout(panel)
        panel_layout.setSpacing(10)
        panel_layout.setContentsMargins(8, 15, 8, 8)

        # Create measurement table with scroll area
        table_scroll = QScrollArea()
        table_scroll.setWidgetResizable(True)
        table_scroll.setMinimumHeight(200)

        measurement_table = QTableWidget(1 if zone_name == "Zone1-Inner" else len(self._relay_values), 3)
        measurement_table.setHorizontalHeaderLabels(["Setpt", "R", "P/F"])
        measurement_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        measurement_table.verticalHeader().setVisible(False)
        measurement_table.setFont(QFont('Arial', 9))

        # Style the table
        measurement_table.setStyleSheet("""
               QTableWidget {
                   background-color: white;
                   gridline-color: #e0e0e0;
               }
               QHeaderView::section {
                   background-color: #e0e0e0;
                   padding: 4px;
                   border: 1px solid #e0e0e0;
                   font-weight: bold;
               }
               QTableWidget::item {
                   background-color: #f5f5f5;
               }
           """)

        # Populate setpoint values
        if zone_name == "Zone1-Inner":
            measurement_table.setItem(0, 0, QTableWidgetItem(str(self._relay_values[0])))
        else:
            for row, relay in enumerate(self._relay_values):
                measurement_table.setItem(row, 0, QTableWidgetItem(str(relay)))

        table_scroll.setWidget(measurement_table)
        panel_layout.addWidget(table_scroll)

        # Add test button
        test_button = QPushButton(f"Test {zone_name}")
        test_button.setFont(QFont('Arial', 9))
        test_button.setFixedHeight(30)
        test_button.setStyleSheet("""
               QPushButton {
                   min-width: 80px;
                   padding: 4px;
                   background: #e0e0e0;
                   border: 1px solid #aaa;
               }
               QPushButton:hover {
                   background: #d0d0d0;
               }
           """)
        test_button.clicked.connect(lambda _, z=zone_name: self._start_resistance_zone_measurement(z))
        panel_layout.addWidget(test_button, alignment=Qt.AlignCenter)

        self._measurement_tables[zone_name] = measurement_table

        return panel

    def _clear_resistance_log_display(self):
        self._log_output.clear()

    def _clear_impedance_log_display(self):
        self._log_output_imp.clear()

    def get_table_item_safe(table, row, column):
        item = table.item(row, column)
        return item.text() if item is not None else ""

    def _get_zone_title(self, zone_name):

        zone_titles = {
            "Zone1-Inner": "Zone 1 - Inner",
            "Zone2-Mid_Inner": "Zone 2 - Mid Inner",
            "Zone3-Mid_Edge": "Zone 3 - Mid Edge",
            "Zone4-Edge": "Zone 4 - Edge",
            "Zone5-Outer": "Zone 5 - Outer"
        }
        return zone_titles.get(zone_name, zone_name)

    def process_single_measurement(self, zone_name, measurement_line):
        """
        Process single measurement line from RPi in format: setpoint,resistance,status
        Example: "0,2.5,True"
        """
        try:
            # Remove any whitespace and split the components
            # setpoint, resistance, status = measurement_line.strip().split(',')
            val = measurement_line.strip().split(',')
            if len(val) < 4:
                self._log_resistance_message(f"Invalid measurement format (expected 4+ fields): {measurement_line}", is_error=True)
                return False
            setpoint = val[1]
            resistance = val[2]
            status = val[3]
            # Find the table for this zone
            table = self._measurement_tables.get(zone_name)
            self.config1 = self.load_config(self.assembly_suffix)

            if not table:
                self._log_resistance_message(f"No table found for zone {zone_name}", is_error=True)
                return False

            # Find the row with matching setpoint
            for row in range(table.rowCount()):
                #self._log_resistance_message(f"Inside block1 ", is_error=True)
                if table.item(row, 0).text() == setpoint:

                    #self._log_resistance_message(f"Inside block2 {table.item(row, 0).text()}", is_error=True)
                    # Convert status to display format
                    status_text = "PASS" if status.lower() == "true" else "FAIL"
                    measurement_data = {
                        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'zone_title': self._get_zone_title(zone_name),
                        'setpoint': float(setpoint),
                        'resistance': float(resistance),
                        'status': "PASS" if status.lower() == "true" else "FAIL",
                        'table_row': row + 1
                    }

                    if status == 'False':
                        self.resistance_test = 'FAIL'
                        self.over_all_result = 'FAIL'

                    self.step_no = self.step_no + 1
                    if float(setpoint) != 0.0:
                        limit_per = float(self.config1[setpoint])
                        prod_val = float(self.config1[f'Res{setpoint}'])
                        higher_limit = prod_val + (prod_val * limit_per / 100)
                        lower_limit = prod_val - (prod_val * limit_per / 100)
                        #self._log_resistance_message(f"Inside block3", is_error=True)
                        # Log metadata along with test data
                        self.excel_logger.log_summary(
                            step_data={
                                'step': str(self.step_no),
                                'unit': 'ohm',
                                'low_limit': f'{lower_limit}',
                                'measure': f'{float(resistance)}',
                                'high_limit': f'{higher_limit}',
                                'teststep': 'Resistance Test',
                                'testpoints': f'{self._get_zone_title(zone_name)}_{setpoint}',
                                'status': "PASS" if status.lower() == "true" else "FAIL"
                            }
                        )
                    else:
                        self.excel_logger.log_summary(
                            step_data={
                                'step': str(self.step_no),
                                'unit': 'ohm',
                                'low_limit': '0.0',
                                'measure': f'{float(resistance)}',
                                'high_limit': '0.7',
                                'teststep': 'Resistance Test',
                                'testpoints': f'{self._get_zone_title(zone_name)}_{setpoint}',
                                'status': "PASS" if status.lower() == "true" else "FAIL"
                            }
                        )

                    #if self._get_zone_title(zone_name) == "Zone 5 - Outer" and float(setpoint) == 127.0:

                    #self._log_resistance_message(f"Inside block4", is_error=True)
                    self.excel_logger.log_summary(
                            metadata={
                                'end_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                'overall_result': self.over_all_result
                            }
                        )
                    self.excel_logger.update_overall_result(self.over_all_result)

                    #self._log_resistance_message(f"Inside block5", is_error=True)
                    # Update the table
                    self.update_resistance_measurement(
                        zone_name,
                        row,
                        float(resistance),
                        status_text
                    )

                    #self._log_resistance_message(f"Inside block6", is_error=True)
                    # Log to Excel (combined sheet)
                    if not self.excel_logger.log_resistance_measurement(measurement_data, f'{zone_name}_Res_scan'):
                        raise Exception("Failed to log to Excel")
                    return True

            self._log_resistance_message(f"Setpoint {setpoint} not found in {zone_name}", is_error=True)
            return False

        except ValueError:
            self._log_resistance_message(f"Invalid measurement format: {measurement_line}", is_error=True)
            return False
        except Exception as e:
            self._log_resistance_message(f"Error processing measurement: {str(e)}", is_error=True)
            return False

    def _start_impedance_zone_measurement(self, zone_name):

        try:

            self.names1 = zone_name
            self.stop_increment = False
            selected_freq = self.freq_combo.currentText()
            frequency = selected_freq.split()[0]
            freq_suffix = ExcelLogger._freq_to_sheet_suffix(selected_freq)
            if selected_freq == "60 MHz":
                command = f"{zone_name} 50000 70000"
            else:
                command = f"{zone_name} {frequency}"


            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(f"Please connect Zone {zone_name} and click OK to continue")
            msg.setWindowTitle("Impedance Scan")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            retval = msg.exec_()
            if retval == QMessageBox.Ok:
                self.start_time = time.time()
                if zone_name == "Zone1-Inner":
                    if self.Zone1_Inner_imp >= 1:
                        excel_logger.reset_sheet(f"Zone1-Inner_{freq_suffix}_Imp_scan")
                        self.stop_increment =  True
                    self.Zone1_Inner_imp += 1

                elif zone_name == "Zone2-Mid_Inner":
                    if self.Zone2_Mid_Inner_imp >= 1:
                        excel_logger.reset_sheet(f"Zone2-Mid_Inner_{freq_suffix}_Imp_scan")
                        self.stop_increment = True
                    self.Zone2_Mid_Inner_imp += 1

                elif zone_name == "Zone3-Mid_Edge":
                    if self.Zone3_Mid_Edge_imp >= 1:
                        excel_logger.reset_sheet(f"Zone3-Mid_Edge_{freq_suffix}_Imp_scan")
                        self.stop_increment = True
                    self.Zone3_Mid_Edge_imp += 1

                elif zone_name == "Zone4-Edge":
                    if self.Zone4_Edge_imp >= 1:
                        excel_logger.reset_sheet(f"Zone4-Edge_{freq_suffix}_Imp_scan")
                        self.stop_increment = True
                    self.Zone4_Edge_imp += 1

                elif zone_name == "Zone5-Outer":
                    if self.Zone5_Outer_imp >= 1:
                        excel_logger.reset_sheet(f"Zone5-Outer_{freq_suffix}_Imp_scan")
                        self.stop_increment = True
                    self.Zone5_Outer_imp += 1

                else:
                    pass

                if hasattr(self, 'worker') and self.worker:
                    self.worker.stop()
                if selected_freq == "60 MHz":
                    script_path = '/home/robot/Manufacturing_test/aipc_beta/VNA_start_stop_60Mhz.py'
                else:
                    script_path = '/home/robot/Manufacturing_test/aipc_beta/VNA_Final.py'
                self.worker = Worker(
                    self.ssh_handler,
                    script_path,
                    command
                )








                self.worker.output_ready.connect(self.handle_Zone_impedance_output)
                self.worker.error_occurred.connect(self.handle_imp_error)
                # Start the thread
                self.worker.start()
                self._log_Impedance_message(f"Starting measurement for {zone_name}")
            else:
                self._log_Impedance_message(f"Impedance Scan {zone_name} suspended,is_error= True")


        except Exception as e:
            self._log_Impedance_message(f"Failed to start test: {str(e)}", is_error=True)
            self.cleanup_resources()


    def handle_imp_error(self, error_msg):
        self._log_Impedance_message(f"ERROR: {error_msg}", is_error=True)
        self.cleanup_resources()

    def handle_config_test(self):
        if hasattr(self, 'worker') and self.worker:
            self.worker.stop()

        # Raspberry Pi SSH details
        #host = "192.168.1.2"  # Replace with your Pi's IP
        host = "10.119.9.225"
        #host =
        username = "robot"
        password = "robot"  # Default, change if needed

        # File paths
        local_file = "C:\\Config\\config.json"  # Windows path
        remote_file = "/home/robot/Manufacturing_test/aipc_beta/config.json"  # Destination on Pi

        # Initialize SSH client
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(host, username=username, password=password)

        # Transfer file using SFTP
        sftp = ssh.open_sftp()
        sftp.put(local_file, remote_file)
        sftp.close()
        ssh.close()


    def file_transer(self,local_file,remote_file):
        if hasattr(self, 'worker') and self.worker:
            self.worker.stop()

        # Raspberry Pi SSH details
        #host = "192.168.1.2"  # Replace with your Pi's IP
        host = "10.119.9.225"
        username = "robot"
        password = "robot"  # Default, change if needed
        self.append_console_message("Transferring the APPOTP file to RPI")
        # File paths
        #local_file = "C:\\Config\\config.json"  # Windows path
        #remote_file = "/home/robot/Manufacturing_test/aipc_beta/config.json"  # Destination on Pi

        # Initialize SSH client
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(host, username=username, password=password)

        # Transfer file using SFTP
        sftp = ssh.open_sftp()
        sftp.put(local_file, remote_file)
        sftp.close()
        ssh.close()



    def process_single_imp_measurement(self, zone_name, measurement_line):
        """
        Process single measurement line from RPi in format: setpoint,resistance,status
        Example: "0,2.5,True"
        """
        try:
            # Remove any whitespace and split the components
            # setpoint, resistance, status = measurement_line.strip().split(',')
            val = measurement_line.strip().split(',')
            if len(val) < 6:
                self._log_Impedance_message(f"Invalid measurement format (expected 6+ fields): {measurement_line}", is_error=True)
                return False
            setpoint = val[1]
            real = val[2]
            img = val[3]
            imped = val[4]
            status = val[5]
            # Find the table for this zone
            table = self._measurement_tables_imp.get(zone_name)
            self.config2 = self.load_config(self.assembly_suffix)

            if not table:
                self._log_Impedance_message(f"No table found for zone {zone_name}", is_error=True)
                return False

            # Find the row with matching setpoint

            for row in range(table.rowCount()):
                if table.item(row, 0).text() == setpoint:
                    # Convert status to display format
                    status_text = "PASS" if status.lower() == "true" else "FAIL"
                    measurement_data = {
                        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'zone_title': zone_name,
                        'Frequency': self.freq_combo.currentText(),
                        'setpoint': setpoint,
                        'Real': float(real),
                        'Imag': float(img),
                        'Z': float(imped),
                        'status': "PASS" if status.lower() == "true" else "FAIL",
                        'table_row': row + 1
                    }


                    if status == 'False':
                        self.Impedance_test = 'FAIL'
                        self.over_all_result = 'FAIL'

                    self.step_no = self.step_no + 1
                    freq_suffix = ExcelLogger._freq_to_sheet_suffix(self.freq_combo.currentText())
                    if float(setpoint) != 0.0:
                        limit_per = float(self.config2[setpoint])
                        prod_val = float(self.config2[zone_name][setpoint])
                        higher_limit = prod_val + (prod_val * limit_per / 100)
                        lower_limit = prod_val - (prod_val * limit_per / 100)

                        # Log metadata along with test data
                        step_data_dict = {
                            'step': str(self.step_no),
                            'unit': 'ohm',
                            'low_limit': f'{lower_limit}',
                            'measure': f'{float(imped)}',
                            'high_limit': f'{higher_limit}',
                            'teststep': 'Impedance Test',
                            'testpoints': f'{zone_name}_{setpoint}_{self.freq_combo.currentText()}',
                            'status': "PASS" if status.lower() == "true" else "FAIL"
                        }
                        self.excel_logger.log_summary(step_data=step_data_dict)
                    else:
                        step_data_dict = {
                            'step': str(self.step_no),
                            'unit': 'ohm',
                            'low_limit': '0.0',
                            'measure': f'{float(imped)}',
                            'high_limit': '0.7',
                            'teststep': 'Impedance Test',
                            'testpoints': f'{zone_name}_{setpoint}_{self.freq_combo.currentText()}',
                            'status': "PASS" if status.lower() == "true" else "FAIL"
                        }
                        self.excel_logger.log_summary(step_data=step_data_dict)

                    #if zone_name == "Zone5-Outer" and float(setpoint) == 143.0:
                    """
                    self.excel_logger.log_summary(
                            teststep_data={
                                'teststep': 'Impedance Test',
                                'status': self.Impedance_test
                            })

                    self.excel_logger.log_summary(
                            metadata={
                                'end_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                'overall_result': self.over_all_result
                            }
                        )
                    """
                    self.excel_logger.update_overall_result(self.over_all_result)
                    # Update the table
                    self.update_impedance_measurement(
                        zone_name,
                        row,
                        float(real),
                        float(img),
                        float(imped),
                        status_text
                    )
                    # Log to Excel (frequency-specific combined sheet)
                    if not self.excel_logger.log_Imp_measurement(measurement_data, f'{zone_name}_{freq_suffix}_Imp_scan'):
                        raise Exception("Failed to log to Excel")
                    return True

            self._log_Impedance_message(f"Setpoint {setpoint} not found in {zone_name}", is_error=True)
            return False

        except ValueError:
            self._log_Impedance_message(f"Invalid measurement format: {measurement_line}", is_error=True)
            return False
        except Exception as e:
            self._log_Impedance_message(f"Error processing measurement: {str(e)}", is_error=True)
            return False

    def _log_Impedance_message(self, message, is_error=False):
        if is_error:
            self._log_output_imp.append(f'<span style="color:red; font-weight:bold;">{message}</span>')
        else:
            self._log_output_imp.append(f'<span style="color:green; font-weight:bold;">{message}</span>')

        self._log_output_imp.verticalScrollBar().setValue(
            self._log_output_imp.verticalScrollBar().maximum()
        )

    def handle_Zone_impedance_output(self, line):
        # self._log_resistance_message(f"{line}")
        # self._log_Impedance_message(f"{line}")
        # self._log_Impedance_message(self.names1)
        if self.names1 in line:
            self._log_Impedance_message(f"Starting measurement for {line}")

            # Process the single measurement line
            if line.strip():
                if not self.process_single_imp_measurement(self.names1, line):
                    self._log_Impedance_message("Failed to process measurement", is_error=True)
            else:
                self._log_Impedance_message("No measurement data received", is_error=True)

        if "Error in slave initialization" in line:
            QMessageBox.critical(
                self,
                "Critical Error",
                "Slave initialization failed! Please check the EtherCAT connection and restart the test."
            )
            self.worker.stop()

        if "no ping" in line:
            self._log_Impedance_message(f"VNA not connected to the Network", is_error= True)
            self.worker.stop()

        if "Test_done" in line:
            self._log_Impedance_message(f"============{self.names1} test completed =====")
            self.worker.stop()
        if time.time() - self.start_time > 300:
            self._log_Impedance_message(
                f"===No Data from Raspberry pi for more than 30 sec please check the Raspberry pi =====", is_error=True)
            self.worker.stop()

    def update_impedance_measurement(self, zone_name, row_index, real, imag, Imp, status):
        """Update a single measurement in the table"""
        if zone_name in self._measurement_tables_imp:
            table = self._measurement_tables_imp[zone_name]
            if 0 <= row_index < table.rowCount():
                # Update resistance value (3 decimal places)
                real_item = QTableWidgetItem(f"{real:.2f}")
                real_item.setFlags(real_item.flags() ^ Qt.ItemIsEditable)
                table.setItem(row_index, 1, real_item)

                image_item = QTableWidgetItem(f"{imag:.2f}")
                image_item.setFlags(image_item.flags() ^ Qt.ItemIsEditable)
                table.setItem(row_index, 2, image_item)

                Imp_item = QTableWidgetItem(f"{Imp:.2f}")
                Imp_item.setFlags(Imp_item.flags() ^ Qt.ItemIsEditable)
                table.setItem(row_index, 3, Imp_item)

                # Update status with appropriate coloring
                status_item = QTableWidgetItem(status)
                status_item.setFlags(status_item.flags() ^ Qt.ItemIsEditable)

                if status == "PASS":
                    status_item.setBackground(QColor(220, 255, 220))  # Light green
                    status_item.setForeground(QColor(0, 128, 0))  # Dark green text
                else:
                    status_item.setBackground(QColor(255, 220, 220))  # Light red
                    status_item.setForeground(QColor(139, 0, 0))  # Dark red text

                table.setItem(row_index, 4, status_item)

                # Scroll to show the updated row
                table.scrollToItem(table.item(row_index, 0))

    def _start_resistance_zone_measurement(self, zone_name):
        try:
            self.names = zone_name
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText(f"Please connect Zone {zone_name} and click OK to continue")
            msg.setWindowTitle("Resistance Scan")
            msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
            retval = msg.exec_()
            if retval == QMessageBox.Ok:
                self.start_time = time.time()
                if zone_name == "Zone1-Inner":
                    if self.Zone1_Inner_res >= 1:
                        excel_logger.reset_sheet("Zone1-Inner_Res_scan")
                    self.Zone1_Inner_res += 1

                elif zone_name == "Zone2-Mid_Inner":
                    if self.Zone2_Mid_Inner_res >= 1:
                        excel_logger.reset_sheet("Zone2-Mid_Inner_Res_scan")
                    self.Zone2_Mid_Inner_res += 1

                elif zone_name == "Zone3-Mid_Edge":
                    if self.Zone3_Mid_Edge_res >= 1:
                        excel_logger.reset_sheet("Zone3-Mid_Edge_Res_scan")
                    self.Zone3_Mid_Edge_res += 1

                elif zone_name == "Zone4-Edge":
                    if self.Zone4_Edge_res >= 1:
                        excel_logger.reset_sheet("Zone4-Edge_Res_scan")
                    self.Zone4_Edge_res += 1

                elif zone_name == "Zone5-Outer":
                    if self.Zone5_Outer_res >= 1:
                        excel_logger.reset_sheet("Zone5-Outer_Res_scan")
                    self.Zone5_Outer_res += 1

                else:
                    pass
                if hasattr(self, 'worker') and self.worker:
                    self.worker.stop()
                self.worker = Worker(
                    self.ssh_handler,
                    '/home/robot/Manufacturing_test/aipc_beta/resistance_test.py',
                    zone_name
                )

                self.worker.output_ready.connect(self.handle_Zone_output)
                self.worker.error_occurred.connect(self.handle_res_error)
                # Start the thread
                self.worker.start()
                self._log_resistance_message(f"Starting measurement for {zone_name}")

            else:
                self._log_resistance_message(f"Resistance Scan : {zone_name} suspended", is_error=True)

        except Exception as e:
            self._log_resistance_message(f"Failed to start test: {str(e)}", is_error=True)
            self.cleanup_resources()


    def handle_res_error(self, error_msg):
        self._log_resistance_message(f"ERROR: {error_msg}", is_error=True)
        self.cleanup_resources()


    def handle_Zone_output(self, line):
        # self._log_resistance_message(f"{line}")
        if "no ping" in line:
            self._log_Impedance_message(f"DIMM not connected to the Network", is_error=True)
            self.worker.stop()

        if "Error in slave initialization" in line:
            QMessageBox.critical(
                self,
                "Critical Error",
                "Slave initialization failed! Please check the EtherCAT connection and restart the test."
            )
            self.worker.stop()

        if self.names in line:
            self._log_resistance_message(f"Starting measurement for {line}")

            # Process the single measurement line
            if line.strip():
                if not self.process_single_measurement(self.names, line):
                    self._log_resistance_message("Failed to process measurement", is_error=True)
            else:
                self._log_resistance_message("No measurement data received", is_error=True)
        if "Test_done" in line:
            self._log_resistance_message(f"============{self.names} test completed =====")
            self.worker.stop()
        if time.time() - self.start_time > 150:
            self._log_resistance_message(
                f"===No Data from Raspberry pi for more than 90 sec please check the Raspberry pi =====", is_error=True)
            self.worker.stop()


    def update_resistance_measurement(self, zone_name, row_index, resistance_value, status):
        """Update a single measurement in the table"""
        if zone_name in self._measurement_tables:
            table = self._measurement_tables[zone_name]
            if 0 <= row_index < table.rowCount():
                # Update resistance value (3 decimal places)
                resistance_item = QTableWidgetItem(f"{resistance_value:.3f}")
                resistance_item.setFlags(resistance_item.flags() ^ Qt.ItemIsEditable)
                table.setItem(row_index, 1, resistance_item)

                # Update status with appropriate coloring
                status_item = QTableWidgetItem(status)
                status_item.setFlags(status_item.flags() ^ Qt.ItemIsEditable)

                if status == "PASS":
                    status_item.setBackground(QColor(220, 255, 220))  # Light green
                    status_item.setForeground(QColor(0, 128, 0))  # Dark green text
                else:
                    status_item.setBackground(QColor(255, 220, 220))  # Light red
                    status_item.setForeground(QColor(139, 0, 0))  # Dark red text

                table.setItem(row_index, 2, status_item)

                # Scroll to show the updated row
                table.scrollToItem(table.item(row_index, 0))


    def _log_resistance_message(self, message, is_error=False):
        if is_error:
            self._log_output.append(f'<span style="color:red; font-weight:bold;">{message}</span>')
        else:
            self._log_output.append(f'<span style="color:green; font-weight:bold;">{message}</span>')
        self._log_output.verticalScrollBar().setValue(
            self._log_output.verticalScrollBar().maximum()
        )


    def create_test_tab(self, title, show_progress=False):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        test_section = QGroupBox()
        test_section.setStyleSheet("""
                QGroupBox {
                    background-color: white;
                    padding: 15px;
                    border-radius: 5px;
                }
                QGroupBox::title {
                    font-weight: bold;
                    font-size: 16px;
                    subcontrol-origin: margin;
                    left: 3px;
                    padding: 0 3px;
                }
            """)
        test_layout = QVBoxLayout(test_section)
        controls = QWidget()
        controls_layout = QHBoxLayout(controls)
        controls_layout.setContentsMargins(0, 0, 0, 0)

        if title == "Interlock System Check":
            # Create the interlock test interface with console-like output

            # Status indicators
            status_layout = QHBoxLayout()
            self.test_status_label = QLabel('Test: Ready')
            self.test_status_label.setAlignment(Qt.AlignCenter)
            self.test_status_label.setStyleSheet("""
                    QLabel {
                        background-color: #17a2b8;
                        color: white;
                        padding: 2px 5px;
                        border-radius: 3px;
                        font-weight: bold;
                        font-size: 9pt;
                    }
                """)
            status_layout.addWidget(self.test_status_label)

            test_layout.addLayout(status_layout)

            # Test controls
            controls_layout = QHBoxLayout()

            self.interlock_start_button = QPushButton('Start Test')
            self.interlock_start_button.setStyleSheet("background-color: #28a745; color: white;")
            self.interlock_start_button.clicked.connect(self.start_interlock_test)
            controls_layout.addWidget(self.interlock_start_button)

            self.interlock_end_button = QPushButton('End Test')
            self.interlock_end_button.setStyleSheet("background-color: #dc3545; color: white;")
            self.interlock_end_button.clicked.connect(self.end_interlock_test)
            self.interlock_end_button.setEnabled(False)
            controls_layout.addWidget(self.interlock_end_button)

            test_layout.addLayout(controls_layout)

            # Interlock status indicators
            self.interlock_status_layout = QHBoxLayout()

            self.interlock_open_label = QLabel('OPEN')
            self.interlock_open_label.setStyleSheet("""
                    QLabel {
                        background-color: lightgray;
                        color: black;
                        padding: 5px;
                        border-radius: 3px;
                        font-weight: bold;
                    }
                """)
            self.interlock_open_label.setAlignment(Qt.AlignCenter)
            self.interlock_status_layout.addWidget(self.interlock_open_label)

            self.interlock_closed_label = QLabel('CLOSED')
            self.interlock_closed_label.setStyleSheet("""
                    QLabel {
                        background-color: lightgray;
                        color: black;
                        padding: 5px;
                        border-radius: 3px;
                        font-weight: bold;
                    }
                """)
            self.interlock_closed_label.setAlignment(Qt.AlignCenter)
            self.interlock_status_layout.addWidget(self.interlock_closed_label)

            test_layout.addLayout(self.interlock_status_layout)
            self.interlock_console = QTextBrowser()
            self.interlock_console.setMinimumHeight(1000)  # Reduced size
            self.interlock_console.setMaximumHeight(1500)
            # self.interlock_console.setMaximumBlockCount(500)
            # self.interlock_console.document().setMaximumBlockCount(500)
            self.interlock_console.setStyleSheet("""
                                        QTextBrowser {
                                            background-color: #f5f5f5;
                                            border: 1px solid #ddd;
                                            font-family: monospace;
                                            font-size: 10pt;
                                        }
                                    """)
            test_layout.addWidget(self.interlock_console)

        elif title == "System Self Test":
            status_layout = QHBoxLayout()
            self.test_status_label_start = QLabel('Test: Ready')
            self.test_status_label_start.setAlignment(Qt.AlignCenter)
            self.test_status_label_start.setStyleSheet("""
                                QLabel {
                                    background-color: #17a2b8;
                                    color: white;
                                    padding: 2px 5px;
                                    border-radius: 3px;
                                    font-weight: bold;
                                    font-size: 9pt;
                                }
                            """)
            status_layout.addWidget(self.test_status_label_start)

            test_layout.addLayout(status_layout)
            controls_layout = QHBoxLayout()

            self.self_start_button = QPushButton('Start Test')
            self.self_start_button.setStyleSheet("background-color: #28a745; color: white;")
            self.self_start_button.clicked.connect(self.start_self_test)
            # self.self_start_button.clicked.connect()
            controls_layout.addWidget(self.self_start_button)
            test_layout.addLayout(controls_layout)

            self.selftest_console = QTextBrowser()
            self.selftest_console.setMinimumHeight(500)  # Reduced size
            self.selftest_console.setMaximumHeight(1500)
            # self.interlock_console.setMaximumBlockCount(500)
            # self.interlock_console.document().setMaximumBlockCount(500)
            self.selftest_console.setStyleSheet("""
                                                    QTextBrowser {
                                                        background-color: #f5f5f5;
                                                        border: 1px solid #ddd;
                                                        font-family: monospace;
                                                        font-size: 10pt;
                                                    }
                                                """)
            test_layout.addWidget(self.selftest_console)

        elif title == "Impedance Scan":
            # Create the main widget for this tab
            impedance_widget = QWidget()
            impedance_layout = QVBoxLayout(impedance_widget)
            impedance_layout.setContentsMargins(5, 5, 5, 5)
            impedance_layout.setSpacing(10)

            # Frequency selection controls
            control_panel = QGroupBox("Test Parameters")
            control_panel.setStyleSheet("""
                       QGroupBox {
                           background-color: #f0f0f0;
                           border: 1px solid #d0d0d0;
                           border-radius: 4px;
                           margin-top: 10px;
                           padding-top: 15px;
                       }
                       QLabel {
                           font-weight: bold;
                           color: #333333;
                       }
                       QComboBox {
                           background-color: white;
                           border: 1px solid #d0d0d0;
                           padding: 3px;
                           font-weight: bold;
                           color: #222222;
                           min-width: 100px;
                       }
                   """)

            control_layout = QHBoxLayout(control_panel)

            # Frequency selection
            self.frequencies = ["362.3 KHz", "400 KHz", "500 KHz", "50 MHz", "60 MHz", "70 MHz"]
            freq_label = QLabel("Test Frequency:")
            freq_label.setFont(QFont('Arial', 10))
            self.freq_combo = QComboBox()
            self.freq_combo.addItems(self.frequencies)
            self.freq_combo.setCurrentIndex(0)  # Default selection
            self.freq_combo.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            control_layout.addWidget(freq_label)
            control_layout.addWidget(self.freq_combo)

            control_layout.addStretch()

            impedance_layout.addWidget(control_panel)

            # Zone setup
            self._zone_names_imp = ["Zone1-Inner", "Zone2-Mid_Inner", "Zone3-Mid_Edge", "Zone4-Edge", "Zone5-Outer"]
            self._relay_values_imp = [0, 1, 2, 4, 8, 16, 32, 64, 127, 128, 135, 141, 142, 143]
            self._measurement_tables_imp = {}

            # Create a scroll area for the zones
            zones_scroll = QScrollArea()
            zones_scroll.setWidgetResizable(True)
            zones_container = QWidget()
            zones_layout = QHBoxLayout(zones_container)
            zones_layout.setContentsMargins(5, 5, 5, 5)
            zones_layout.setSpacing(10)

            # Create zone panels with minimum sizing
            for zone in self._zone_names_imp:
                zone_panel = self._create_Impedance_zone_panel(zone)
                zone_panel.setMinimumWidth(250)  # Minimum width for each zone
                zones_layout.addWidget(zone_panel)

            zones_scroll.setWidget(zones_container)
            impedance_layout.addWidget(zones_scroll, stretch=1)  # Takes most of the space

            # Console output at the bottom
            console_group = QGroupBox("Measurement Log")
            console_layout = QVBoxLayout(console_group)
            self._log_output_imp = QTextBrowser()
            self._log_output_imp.setMinimumHeight(150)
            self._log_output_imp.setStyleSheet("""
                       QTextBrowser {
                           background-color: #f5f5f5;
                           border: 1px solid #ddd;
                           font-family: monospace;
                           font-size: 10pt;
                       }
                   """)

            clear_button = QPushButton("Clear Log")
            clear_button.clicked.connect(self._clear_impedance_log_display)
            clear_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            console_layout.addWidget(self._log_output_imp)
            console_layout.addWidget(clear_button, alignment=Qt.AlignRight)

            impedance_layout.addWidget(console_group, stretch=0)  # Takes less space

            test_layout.addWidget(impedance_widget)

        elif title == "Resistance Test":
            # Create the main widget for this tab
            resistance_widget = QWidget()
            resistance_layout = QVBoxLayout(resistance_widget)
            resistance_layout.setContentsMargins(5, 5, 5, 5)
            resistance_layout.setSpacing(10)

            # Zone names and relay values
            self._zone_names = ["Zone1-Inner", "Zone2-Mid_Inner", "Zone3-Mid_Edge", "Zone4-Edge", "Zone5-Outer"]
            self._relay_values = [0, 1, 2, 4, 8, 16, 32, 64, 127]
            self._measurement_tables = {}

            # Create a scroll area for the zones
            zones_scroll = QScrollArea()
            zones_scroll.setWidgetResizable(True)
            zones_container = QWidget()
            zones_layout = QHBoxLayout(zones_container)
            zones_layout.setContentsMargins(5, 5, 5, 5)
            zones_layout.setSpacing(10)

            # Create zone panels with minimum sizing
            for zone in self._zone_names:
                zone_panel = self._create_resistance_zone_panel(zone)
                zone_panel.setMinimumWidth(250)  # Minimum width for each zone
                zones_layout.addWidget(zone_panel)

            zones_scroll.setWidget(zones_container)
            resistance_layout.addWidget(zones_scroll, stretch=1)  # Takes most of the space

            # Console output at the bottom
            console_group = QGroupBox("Measurement Log")
            console_layout = QVBoxLayout(console_group)
            self._log_output = QTextBrowser()
            self._log_output.setMinimumHeight(150)
            self._log_output.setStyleSheet("""
                        QTextBrowser {
                            background-color: #f5f5f5;
                            border: 1px solid #ddd;
                            font-family: monospace;
                            font-size: 10pt;
                        }
                    """)

            clear_button = QPushButton("Clear Log")
            clear_button.clicked.connect(self._clear_resistance_log_display)
            clear_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            console_layout.addWidget(self._log_output)
            console_layout.addWidget(clear_button, alignment=Qt.AlignRight)

            resistance_layout.addWidget(console_group, stretch=0)  # Takes less space

            test_layout.addWidget(resistance_widget)

        elif title == "DIMM Calibration":
            DIMM_layout = QHBoxLayout()
            self.DIMM_status_label_start = QLabel('Test: Ready')
            self.DIMM_status_label_start.setAlignment(Qt.AlignCenter)
            self.DIMM_status_label_start.setStyleSheet("""
                                         QLabel {
                                             background-color: #17a2b8;
                                             color: white;
                                             padding: 2px 5px;
                                             border-radius: 3px;
                                             font-weight: bold;
                                             font-size: 9pt;
                                         }
                                     """)
            DIMM_layout.addWidget(self.DIMM_status_label_start)

            test_layout.addLayout(DIMM_layout)
            controls_layout = QHBoxLayout()

            self.dimm_start_button = QPushButton('Start calibration')
            self.dimm_start_button.setStyleSheet("background-color: #28a745; color: white;")
            self.dimm_start_button.clicked.connect(self.dimm_cal_test)
            # self.self_start_button.clicked.connect()
            controls_layout.addWidget(self.dimm_start_button)
            test_layout.addLayout(controls_layout)
            self.dimm_progress = QProgressBar()
            self.dimm_progress.setRange(0, 100)
            self.dimm_progress.setValue(0)
            self.dimm_progress.setMinimumHeight(20)
            self.dimm_progress.setStyleSheet("""
                        QProgressBar {
                            border: 1px solid grey;
                            border-radius: 5px;
                            text-align: center;
                        }
                        QProgressBar::chunk {
                            background-color: #28a745;
                            width: 10px;
                        }
                    """)
            test_layout.addWidget(self.dimm_progress)
            # dimm = QProgressBar()
            # dimm.setValue(0)
            # dimm.setMinimumHeight(20)
            # dimm.addWidget(dimm)

            self.dimmtest_console = QTextBrowser()
            self.dimmtest_console.setMinimumHeight(500)  # Reduced size
            self.dimmtest_console.setMaximumHeight(1500)
            # self.interlock_console.setMaximumBlockCount(500)
            # self.interlock_console.document().setMaximumBlockCount(500)
            self.dimmtest_console.setStyleSheet("""
                                                             QTextBrowser {
                                                                 background-color: #f5f5f5;
                                                                 border: 1px solid #ddd;
                                                                 font-family: monospace;
                                                                 font-size: 10pt;
                                                             }
                                                         """)
            test_layout.addWidget(self.dimmtest_console)
        elif title == "VNA Calibration":
            VNA_layout = QHBoxLayout()
            self.VNA_status_label_start = QLabel('Test: Ready')
            self.VNA_status_label_start.setAlignment(Qt.AlignCenter)
            self.VNA_status_label_start.setStyleSheet("""
                                                     QLabel {
                                                         background-color: #17a2b8;
                                                         color: white;
                                                         padding: 2px 5px;
                                                         border-radius: 3px;
                                                         font-weight: bold;
                                                         font-size: 9pt;
                                                     }
                                                 """)
            VNA_layout.addWidget(self.VNA_status_label_start)

            test_layout.addLayout(VNA_layout)

            controls_layout = QHBoxLayout()

            self.VNA_start_button = QPushButton('Start calibration')
            self.VNA_start_button.setStyleSheet("background-color: #28a745; color: white;")
            self.VNA_start_button.clicked.connect(self.VNA_cal_test)
            # self.self_start_button.clicked.connect()
            controls_layout.addWidget(self.VNA_start_button)
            test_layout.addLayout(controls_layout)
            self.vna_progress = QProgressBar()
            self.vna_progress.setRange(0, 100)
            self.vna_progress.setValue(0)
            self.vna_progress.setMinimumHeight(20)
            self.vna_progress.setStyleSheet("""
                             QProgressBar {
                                 border: 1px solid grey;
                                 border-radius: 5px;
                                 text-align: center;
                             }
                             QProgressBar::chunk {
                                 background-color: #28a745;
                                 width: 10px;
                             }
                         """)
            test_layout.addWidget(self.vna_progress)
            self.VNAtest_console = QTextBrowser()
            self.VNAtest_console.setMinimumHeight(500)  # Reduced size
            self.VNAtest_console.setMaximumHeight(1500)
            # self.interlock_console.setMaximumBlockCount(500)
            # self.interlock_console.document().setMaximumBlockCount(500)
            self.VNAtest_console.setStyleSheet("""
                                                    QTextBrowser {
                                                    background-color: #f5f5f5;
                                                    border: 1px solid #ddd;
                                                    font-family: monospace;
                                                    font-size: 10pt;
                                                    }
                                                    """)
            test_layout.addWidget(self.VNAtest_console)

        elif title == "Verify BNC Port":
            BNC_layout = QHBoxLayout()
            self.BNC_status_label_start = QLabel('Test: Ready')
            self.BNC_status_label_start.setAlignment(Qt.AlignCenter)
            self.BNC_status_label_start.setStyleSheet("""
                                                                 QLabel {
                                                                     background-color: #17a2b8;
                                                                     color: white;
                                                                     padding: 2px 5px;
                                                                     border-radius: 3px;
                                                                     font-weight: bold;
                                                                     font-size: 9pt;
                                                                 }
                                                             """)
            BNC_layout.addWidget(self.BNC_status_label_start)

            test_layout.addLayout(BNC_layout)
            controls_layout = QHBoxLayout()

            self.BNC_start_button = QPushButton('Start')
            self.BNC_start_button.setStyleSheet("background-color: #28a745; color: white;")
            self.BNC_start_button.clicked.connect(self.BNC_test)
            # self.self_start_button.clicked.connect()
            controls_layout.addWidget(self.BNC_start_button)
            test_layout.addLayout(controls_layout)
            if show_progress:
                progress = QProgressBar()
                progress.setValue(0)
                progress.setMinimumHeight(20)
                test_layout.addWidget(progress)
            # BNC = QProgressBar()
            # BNC.setValue(0)
            # BNC.setMinimumHeight(20)
            # BNC.addWidget(BNC)

            self.BNCtest_console = QTextBrowser()
            self.BNCtest_console.setMinimumHeight(500)  # Reduced size
            self.BNCtest_console.setMaximumHeight(1500)
            # self.interlock_console.setMaximumBlockCount(500)
            # self.interlock_console.document().setMaximumBlockCount(500)
            self.BNCtest_console.setStyleSheet("""
                                                    QTextBrowser {
                                                    background-color: #f5f5f5;
                                                    border: 1px solid #ddd;
                                                    font-family: monospace;
                                                    font-size: 10pt;
                                                    }
                                                    """)
            test_layout.addWidget(self.BNCtest_console)


        else:
            pass

        layout.addWidget(test_section)
        return tab


    def append_interlock_message(self, message, is_error=False):
        """Helper method to append colored messages to interlock console"""
        if is_error:
            self.interlock_console.append(f'<span style="color:red; font-weight:bold;">{message}</span>')
        else:
            self.interlock_console.append(f'<span style="color:green;font-weight:bold;">{message}</span>')
        # Auto-scroll to bottom
        self.interlock_console.verticalScrollBar().setValue(
            self.interlock_console.verticalScrollBar().maximum()
        )


    def append_self_message(self, message, is_error=False):
        """Helper method to append colored messages to interlock console"""
        if is_error:
            self.selftest_console.append(f'<span style="color:red; font-weight:bold;">{message}</span>')
        else:
            self.selftest_console.append(f'<span style="color:green;font-weight:bold;">{message}</span>')
        # Auto-scroll to bottom
        self.selftest_console.verticalScrollBar().setValue(
            self.selftest_console.verticalScrollBar().maximum()
        )


    def start_self_test(self):
        try:
            if self.self_t >= 1:
                excel_logger.reset_sheet("Self Test")
            self.self_t += 1
            self.selftest_console.clear()
            success, message = self.ssh_handler.Connect_RPI()
            if not success:
                self.handle_ssh_error(f"Connection failed: {message}")
                self.append_self_message(f"SSH connection Failed", is_error=True)
                return
            self.test_status_label_start.setText("Running")
            self.test_status_label_start.setStyleSheet("background-color: #ffc107; color: black;")
            self.append_self_message("\n==================Self Test Started=======================\n")
            self.append_self_message("\nWait Test in process..........\n")
            self.execute_command("selftest", self.handle_self_test_output, 0)
        except Exception as e:
            self.logger.error(f"Error in self_test : {str(e)}", exc_info=True, extra={'func_name': 'start_self_test'})
            QMessageBox.critical(self, "Error", f"SELF TEST FAIL: {str(e)}")
        finally:
            self.ssh_handler.SSH_disconnect()


    def start_interlock_test(self):
        try:
            # Clean up any existing test
            if self.impedance_scan >= 1:
                excel_logger.reset_sheet("Interlock Test")
            self.impedance_scan += 1
            self.fan_interlock = 30

            if hasattr(self, 'worker') and self.worker:
                try:
                    self.worker.output_ready.disconnect(self.handle_interlock_output)
                    self.worker.finished_signal.disconnect(self.on_interlock_test_finished)
                    self.worker.error_occurred.disconnect(self.handle_interlock_error)
                except (TypeError, RuntimeError):
                    pass
                self.worker.stop()

            # Reset UI state
            self.interlock_console.clear()
            self.append_interlock_message("\n1. Fan Interlock Test Started\n")
            self.open_count = 0
            self.closed_count = 0
            # self.update_interlock_counters()

            # Create new worker
            self.worker = Worker(
                self.ssh_handler,
                '/home/robot/Manufacturing_test/aipc_beta/Interlock.py',
                'ecat test_interlock'
            )

            # Connect signals
            self.worker.output_ready.connect(self.handle_interlock_output)
            self.worker.finished_signal.connect(self.on_interlock_test_finished)
            self.worker.error_occurred.connect(self.handle_interlock_error)

            # Update UI
            self.interlock_start_button.setEnabled(False)
            self.interlock_end_button.setEnabled(True)
            self.test_status_label.setText("Running")
            self.test_status_label.setStyleSheet("background-color: #ffc107; color: black;")

            # Start the thread
            self.worker.start()

        except Exception as e:
            self.append_interlock_message(f"Failed to start test: {str(e)}", is_error=True)
            self.cleanup_resources()


    def handle_interlock_error(self, error_msg):
        self.append_interlock_message(f"ERROR: {error_msg}", is_error=True)
        self.cleanup_resources()


    def handle_interlock_output(self, line):
        """Handle output from the interlock test"""
        if "Error in slave initialization" in line:
            QMessageBox.critical(
                self,
                "Critical Error",
                "Slave initialization failed! Please check the EtherCAT connection and restart the test."
            )
            self.interlock_start_button.setEnabled(True)
            self.worker.stop()

        if "Cooling Fan Working" in line:
            self.fan_interlock = True
            self.append_interlock_message("✔ Fan Interlock Pass \n\n")
            time.sleep(1)
            self.append_interlock_message("\n2. Switch Interlock Test Started\n\n")
            self.append_interlock_message("\nPress the Interlock Switch.......\n")
            time.sleep(1)

        if "Cooling Fan Warning" in line:
            self.fan_interlock = False
            self.over_all_result = "FAIL"
            self.append_interlock_message("✖ Fan Interlock Fail \n\n", is_error=True)
            # time.sleep(1)
            self.append_interlock_message("\n2. Switch Interlock Test Started\n\n")
            self.append_interlock_message("\nPress the Interlock Switch.......\n")
            # time.sleep(1)
            # self.check = True

        if "Interlock Open" in line:
            self.open_count += 1
            self.interlock_open_label.setStyleSheet("background-color: #dc3545; color: white;")
            self.interlock_open_label.setText(f"OPEN")
            if self.open_count == 1:
                self.check_true += 1
                # self.append_interlock_message("Interlock Open detected")
        elif "Interlock Closed" in line:
            self.closed_count += 1
            self.interlock_closed_label.setStyleSheet("background-color: #28a745; color: white;")
            self.interlock_closed_label.setText(f"CLOSED")
            if self.closed_count == 1:
                self.check_true += 1
                self.append_interlock_message("Interlock Closed detected")
        if self.check_true == 1:
            # self.append_interlock_message("Press the Interlock Switch.......")
            # QMessageBox.information("Press the Interlock switch and ok buttom")
            self.check_true += 1


    def end_interlock_test(self):
        try:
            if self.fan_interlock != 30:
                open_count = False
                close_count = False
                if hasattr(self, 'worker') and self.worker:
                    self.worker.stop()

                if self.closed_count >= 1:
                    close_count = True

                if self.open_count >= 1:
                    open_count = True

                test_passed = self.closed_count >= 1 and self.open_count >= 1

                if self.fan_interlock:
                    result_msg_i = "TEST PASSED - Fan Interlock detected properly"
                else:
                    result_msg_i = "TEST FAILED -  Fan Interlock test Fail"

                if test_passed:
                    result_msg = "TEST PASSED - Interlock switch detected properly"
                    self.append_interlock_message(result_msg)
                    self.test_status_label.setText("Passed")
                    self.test_status_label.setStyleSheet("background-color: #28a745; color: white;")
                    count = True
                else:
                    result_msg = f"TEST FAILED -  Interlock test Fail"
                    self.over_all_result = 'FAIL'
                    self.append_interlock_message(result_msg, is_error=True)
                    self.test_status_label.setText("Failed")
                    self.test_status_label.setStyleSheet("background-color: #dc3545; color: white;")
                    # Log to Excel

                self.excel_logger.log_interlock_test(
                    test_name='FAN Interlock',
                    open_count='NA',
                    closed_count='NA',
                    test_passed=self.fan_interlock,
                    notes=result_msg_i
                )

                self.excel_logger.log_summary(
                    teststep_data={
                        'teststep': 'FAN Test',
                        'status': "PASS" if self.fan_interlock else "FAIL"  # Updates the existing row
                    }
                )
                self.excel_logger.log_interlock_test(
                    test_name='Switch Interlock',
                    open_count=open_count,
                    closed_count=close_count,
                    test_passed=test_passed,
                    notes=result_msg
                )
                self.excel_logger.log_summary(
                    teststep_data={
                        'teststep': 'Interlock Test',
                        'status':  "PASS" if test_passed else "FAIL"
                    }
                )

                self.excel_logger.log_summary(
                    metadata={
                        'end_time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        'overall_result': self.over_all_result
                    }
                )
                self.excel_logger.update_overall_result(self.over_all_result)



            else:
                result_msg = f"TEST aborted without any test"
                self.append_interlock_message(result_msg, is_error=True)

        except Exception as e:
            self.append_interlock_message(f"Error ending test: {str(e)}", is_error=True)
            self.worker.stop()
        finally:
            self.interlock_end_button.setEnabled(False)
            self.interlock_start_button.setEnabled(True)


    def cleanup_resources(self):
        try:
            if hasattr(self, 'worker') and self.worker:
                self.worker.stop()

            if hasattr(self, 'ssh_handler') and self.ssh_handler.is_connect:
                self.ssh_handler.SSH_disconnect()

        except Exception as e:
            self.logger.error(f"Cleanup error: {str(e)}", exc_info=True, extra={'func_name': 'cleanup_resources'})


    def on_interlock_test_finished(self):
        """Clean up after test completion"""
        self.ssh_handler.SSH_disconnect()
        # self.ssh_status_label.setText("SSH: Disconnected")
        # self.ssh_status_label.setStyleSheet("background-color: #dc3545; color: white;")


    def reset_interlock_test(self):
        """Reset the test state"""
        self.interlock_start_button.setEnabled(True)
        self.interlock_end_button.setEnabled(False)
        self.test_status_label.setText("Test: Ready")
        self.test_status_label.setStyleSheet("background-color: #17a2b8; color: white;")


    def create_unit_setup_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QHBoxLayout(scroll_content)
        scroll_layout.setContentsMargins(0, 0, 5, 0)
        scroll_layout.setSpacing(10)

        details_section = QGroupBox()
        details_section.setStyleSheet("""
                QGroupBox {
                    background-color: #f0f0f0;
                    padding: 12px;
                    border-radius: 5px;
                    font-size: 16px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 3px;
                    padding: 0 3px;
                }
            """)
        details_layout = QVBoxLayout(details_section)

        # PCB Information Group - Three fields in one line
        pcb_info_group = QWidget()
        pcb_info_layout = QVBoxLayout(pcb_info_group)
        pcb_info_layout.setContentsMargins(0, 0, 0, 0)

        pcb_info_label = QLabel("PCB Information *")
        pcb_info_label.setStyleSheet("color: red;")
        pcb_info_layout.addWidget(pcb_info_label)

        # Horizontal container for the three fields
        pcb_fields_container = QWidget()
        pcb_fields_layout = QHBoxLayout(pcb_fields_container)
        pcb_fields_layout.setContentsMargins(0, 0, 0, 0)
        pcb_fields_layout.setSpacing(10)

        # PCB Part Number (format: XXX-AXXXXX-XXX)
        pcb_pn_group = QWidget()
        pcb_pn_layout = QVBoxLayout(pcb_pn_group)
        pcb_pn_layout.setContentsMargins(0, 0, 0, 0)
        pcb_pn_label = QLabel("Control Part Number")
        pcb_pn_layout.addWidget(pcb_pn_label)

        self.pcb_pn_input = QLineEdit()
        self.pcb_pn_input.setPlaceholderText("Format: 123-A45678-901")
        self.pcb_pn_input.setMinimumHeight(38)
        pcb_pn_layout.addWidget(self.pcb_pn_input)
        pcb_fields_layout.addWidget(pcb_pn_group)

        # PCB Revision (3 characters)
        pcb_rev_group = QWidget()
        pcb_rev_layout = QVBoxLayout(pcb_rev_group)
        pcb_rev_layout.setContentsMargins(0, 0, 0, 0)
        pcb_rev_label = QLabel("Control Revision")
        pcb_rev_layout.addWidget(pcb_rev_label)

        self.pcb_rev_input = QLineEdit()
        self.pcb_rev_input.setPlaceholderText("Format: A")
        self.pcb_rev_input.setMaxLength(3)
        self.pcb_rev_input.setMinimumHeight(38)
        pcb_rev_layout.addWidget(self.pcb_rev_input)
        pcb_fields_layout.addWidget(pcb_rev_group)

        # PCB Serial Number (12 characters)
        pcb_sn_group = QWidget()
        pcb_sn_layout = QVBoxLayout(pcb_sn_group)
        pcb_sn_layout.setContentsMargins(0, 0, 0, 0)
        pcb_sn_label = QLabel("Control Serial Number")
        pcb_sn_layout.addWidget(pcb_sn_label)

        self.pcb_sn_input = QLineEdit()
        # self.pcb_sn_input.setPlaceholderText()
        self.pcb_sn_input.setMaxLength(12)
        self.pcb_sn_input.setMinimumHeight(38)
        pcb_sn_layout.addWidget(self.pcb_sn_input)
        pcb_fields_layout.addWidget(pcb_sn_group)

        pcb_info_layout.addWidget(pcb_fields_container)
        details_layout.addWidget(pcb_info_group)

        # Assembly Information Group - Three fields in one line
        assembly_info_group = QWidget()
        assembly_info_layout = QVBoxLayout(assembly_info_group)
        assembly_info_layout.setContentsMargins(0, 0, 0, 0)

        assembly_info_label = QLabel("Assembly Information *")
        assembly_info_label.setStyleSheet("color: red;")
        assembly_info_layout.addWidget(assembly_info_label)

        # Horizontal container for the three fields
        assembly_fields_container = QWidget()
        assembly_fields_layout = QHBoxLayout(assembly_fields_container)
        assembly_fields_layout.setContentsMargins(0, 0, 0, 0)
        assembly_fields_layout.setSpacing(10)

        # Assembly Part Number (format: XXX-AXXXXX-XXX)
        assembly_pn_group = QWidget()
        assembly_pn_layout = QVBoxLayout(assembly_pn_group)
        assembly_pn_layout.setContentsMargins(0, 0, 0, 0)
        assembly_pn_label = QLabel("Part Number")
        assembly_pn_layout.addWidget(assembly_pn_label)

        self.assembly_pn_input = QLineEdit()
        self.assembly_pn_input.setPlaceholderText("Format: 123-A45678-901")
        self.assembly_pn_input.setMinimumHeight(38)
        assembly_pn_layout.addWidget(self.assembly_pn_input)
        assembly_fields_layout.addWidget(assembly_pn_group)

        # Assembly Revision (3 characters)
        assembly_rev_group = QWidget()
        assembly_rev_layout = QVBoxLayout(assembly_rev_group)
        assembly_rev_layout.setContentsMargins(0, 0, 0, 0)
        assembly_rev_label = QLabel("Revision")
        assembly_rev_layout.addWidget(assembly_rev_label)

        self.assembly_rev_input = QLineEdit()
        self.assembly_rev_input.setPlaceholderText("Format: A")
        self.assembly_rev_input.setMaxLength(3)
        self.assembly_rev_input.setMinimumHeight(38)
        assembly_rev_layout.addWidget(self.assembly_rev_input)
        assembly_fields_layout.addWidget(assembly_rev_group)

        # Assembly Serial Number (12 characters)
        assembly_sn_group = QWidget()
        assembly_sn_layout = QVBoxLayout(assembly_sn_group)
        assembly_sn_layout.setContentsMargins(0, 0, 0, 0)
        assembly_sn_label = QLabel("Serial Number")
        assembly_sn_layout.addWidget(assembly_sn_label)

        self.assembly_sn_input = QLineEdit()
        # self.assembly_sn_input.setPlaceholderText()
        self.assembly_sn_input.setMaxLength(12)
        self.assembly_sn_input.setMinimumHeight(38)
        assembly_sn_layout.addWidget(self.assembly_sn_input)
        assembly_fields_layout.addWidget(assembly_sn_group)

        assembly_info_layout.addWidget(assembly_fields_container)
        details_layout.addWidget(assembly_info_group)

        VN_FN_info_group = QWidget()
        VN_FN_info_layout = QVBoxLayout(VN_FN_info_group)
        VN_FN_info_layout.setContentsMargins(0, 0, 0, 0)

        # Horizontal container for the two fields
        VN_FN_fields_container = QWidget()
        VN_FN_fields_layout = QHBoxLayout(VN_FN_fields_container)
        VN_FN_fields_layout.setContentsMargins(0, 0, 0, 0)
        VN_FN_fields_layout.setSpacing(10)

        # Vendor Name field
        VN_group = QWidget()
        VN_layout = QVBoxLayout(VN_group)  # Fixed: Create new layout for VN_group
        VN_layout.setContentsMargins(0, 0, 0, 0)
        VN_label = QLabel("Vendor Name")
        VN_layout.addWidget(VN_label)

        self.Vendor_name = QLineEdit()
        self.Vendor_name.setMinimumHeight(38)
        VN_layout.addWidget(self.Vendor_name)
        VN_FN_fields_layout.addWidget(VN_group)

        # Fixture Number field
        FN_group = QWidget()
        FN_layout = QVBoxLayout(FN_group)
        FN_layout.setContentsMargins(0, 0, 0, 0)
        FN_label = QLabel("Fixture Number")
        FN_layout.addWidget(FN_label)

        self.Fixture = QLineEdit()
        self.Fixture.setMinimumHeight(38)
        FN_layout.addWidget(self.Fixture)
        VN_FN_fields_layout.addWidget(FN_group)

        VN_FN_info_layout.addWidget(VN_FN_fields_container)
        details_layout.addWidget(VN_FN_info_group)

        # self.Vendor_name = self.create_form_row("Vendor Name", QLineEdit(), details_layout)

        # Rest of the fields
        # self.Test_date = self.create_form_row("Test Date", QLineEdit(), details_layout)
        self.Test_Date = QDateEdit()
        self.Test_Date.setCalendarPopup(True)
        self.Test_Date.setDate(QDate.currentDate())
        self.Test_Date.setDisplayFormat("yyyy-MM-dd")
        self.create_form_row("Test Date", self.Test_Date, details_layout)

        self.Test_name = self.create_form_row("Test Operator Name", QLineEdit(), details_layout)
        # self.VNA_calibration = self.create_form_row("VNA Calibration Date", QLineEdit(), details_layout)
        self.VNA_calibration = QDateEdit()
        self.VNA_calibration.setCalendarPopup(True)
        self.VNA_calibration.setDate(QDate.currentDate())
        self.VNA_calibration.setDisplayFormat("yyyy-MM-dd")
        self.create_form_row("VNA calibration", self.VNA_calibration, details_layout)

        self.VNA_SN = self.create_form_row("VNA SN", QLineEdit(), details_layout)
        self.Ecal_SN = self.create_form_row("Ecal SN", QLineEdit(), details_layout)

        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        details_layout.addWidget(separator)

        self.product_id = self.create_form_row("Product ID", QLineEdit(), details_layout, True)
        self.esi_revision = self.create_form_row("ESI Revision", QLineEdit(), details_layout, True)
        self.configuration_id = self.create_form_row("Configuration ID", QLineEdit(), details_layout, True)
        self.ethercat_address = self.create_form_row("EtherCAT Address", QLineEdit(), details_layout, True)
        self.firmware_version = self.create_form_row("Firmware Version", QLineEdit(), details_layout, True)

        scroll_layout.addWidget(details_section)

        # Configuration section
        config_section = QGroupBox()
        config_section.setStyleSheet(details_section.styleSheet())
        config_layout = QVBoxLayout(config_section)
        config_layout.setSpacing(15)

        def add_config_widget(label, widget, readonly=False):
            container = QWidget()
            container_layout = QVBoxLayout(container)
            container_layout.setContentsMargins(0, 0, 0, 0)
            container_layout.setSpacing(5)

            lbl = QLabel(label)
            container_layout.addWidget(lbl)

            if isinstance(widget, QLineEdit):
                widget.setMinimumHeight(38)
                if readonly:
                    widget.setReadOnly(True)
                    widget.setStyleSheet("""
                            background-color: #e9ecef;
                            color: #495057;
                            border: 1px solid #ced4da;
                            font-weight: bold;
                        """)

            container_layout.addWidget(widget)
            container_layout.addSpacing(10)
            config_layout.addWidget(container)
            return widget

        self.test_purpose = add_config_widget("Test Purpose", QComboBox())
        self.test_purpose.addItem("High Volume Manufacturing Test")

        otp_program_container = QWidget()
        otp_program_layout = QVBoxLayout(otp_program_container)
        otp_program_layout.setContentsMargins(0, 0, 0, 0)
        otp_program_layout.setSpacing(5)

        self.otp_program_btn = QPushButton("OTP Program")
        self.otp_program_btn.setStyleSheet("""
                QPushButton {
                    background-color: #28a745;
                    color: white;
                    padding: 10px;
                    font-size: 30px;
                    font-weight: bold;
                    min-height: 45px;
                }
                QPushButton:hover {
                    background-color: #218838;
                }
                QPushButton:disabled {
                    background-color: #cccccc;
                    color: #666666;
                }
            """)
        self.otp_program_btn.clicked.connect(self.program_otp)  # Connect to your OTP programming function
        otp_program_layout.addWidget(self.otp_program_btn)
        otp_program_layout.addSpacing(15)
        config_layout.addWidget(otp_program_container)

        # self.pcb_pn = add_config_widget("PCB Part Number", QLineEdit(), True)
        # self.pcb_sn = add_config_widget("PCB Serial Number", QLineEdit(), True)
        # self.pcb_revision = add_config_widget("PCB Revision", QLineEdit(), True)
        # self.assembly_pn = add_config_widget("Assembly Part Number", QLineEdit(), True)
        # self.assembly_sn = add_config_widget("Assembly Serial Number", QLineEdit(), True)
        # self.assembly_revision = add_config_widget("Assembly Revision", QLineEdit(), True)

        auto_load_container = QWidget()
        auto_load_layout = QVBoxLayout(auto_load_container)
        auto_load_layout.setContentsMargins(0, 0, 0, 0)
        auto_load_layout.setSpacing(5)

        self.auto_load_btn = QPushButton("Auto Load and Connect")
        self.auto_load_btn.setStyleSheet("""
                QPushButton {
                    background-color: #007bff; 
                    color: white;
                    padding: 10px;
                    font-size: 20 px;
                    font-weight: bold;
                    min-height: 45px;
                }
                QPushButton:hover {
                    background-color: #0069d9;
                }
                QPushButton:disabled {
                    background-color: #cccccc;
                    color: #666666;
                }
            """)
        self.auto_load_btn.clicked.connect(self.auto_load_connect)
        auto_load_layout.addWidget(self.auto_load_btn)
        auto_load_layout.addSpacing(15)
        config_layout.addWidget(auto_load_container)

        console_group = QGroupBox("Console Output")
        console_group.setStyleSheet("""
                QGroupBox {
                    background-color: white;
                    padding: 10px;
                    border-radius: 5px;
                    font-size: 20px;
                    font-weight: bold;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 5px;
                    padding: 0 3px;
                }
            """)
        console_layout = QVBoxLayout(console_group)
        console_layout.setContentsMargins(5, 15, 5, 5)
        console_layout.setSpacing(5)

        console_buttons = QHBoxLayout()
        clear_btn = QPushButton("Clear Console")
        clear_btn.setStyleSheet("""
                QPushButton {
                    background-color: #6c757d;
                    color: white;
                    padding: 5px 10px;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #5a6268;
                }
            """)
        clear_btn.clicked.connect(lambda: self.console_output.clear() if hasattr(self, 'console_output') else None)
        console_buttons.addWidget(clear_btn)
        console_buttons.addStretch()
        console_layout.addLayout(console_buttons)

        self.console_output = QTextEdit()
        self.console_output.setReadOnly(True)
        self.console_output.setAcceptRichText(True)  # Enable rich text formatting
        self.console_output.setStyleSheet("""
                QTextEdit {
                    background-color: #f8f9fa;
                    border: 1px solid #ced4da;
                    border-radius: 4px;
                    font-family: 'Courier New', monospace;
                    font-size: 30px;
                    padding: 5px;
                    min-height: 1000px;
                }
            """)
        console_layout.addWidget(self.console_output)
        config_layout.addWidget(console_group)
        config_layout.addStretch(1)
        # Increase maximum block count for larger output

        scroll_layout.addWidget(config_section)
        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)
        return tab


    def init_ui(self):
        """Initialize the user interface with responsive layouts"""
        # Create a main scroll area
        main_scroll = QScrollArea()
        main_scroll.setWidgetResizable(True)

        # Create the central widget that will hold everything
        central_widget = QWidget()
        self.main_layout = QVBoxLayout(central_widget)

        # Create tab widget
        self.tab_widget = QTabWidget()
        self.main_layout.addWidget(self.tab_widget)

        # Add tabs
        self.tab_widget.addTab(self.create_unit_setup_tab(), "Unit Setup")
        self.tab_widget.addTab(self.create_test_tab("Interlock System Check"), "Interlock")
        self.tab_widget.addTab(self.create_test_tab("Verify BNC Port"), "BNC Port Verification")
        self.tab_widget.addTab(self.create_test_tab("Impedance Scan"), "Impedance Scan")
        self.tab_widget.addTab(self.create_test_tab("Resistance Test"), "Resistance")
        self.tab_widget.addTab(self.create_test_tab("System Self Test"), "Self Test")
        self.tab_widget.addTab(self.create_test_tab("DIMM Calibration"), "DIMM Cal")
        self.tab_widget.addTab(self.create_test_tab("VNA Calibration"), "VNA Cal")

        # Set the central widget
        main_scroll.setWidget(central_widget)
        self.setCentralWidget(main_scroll)


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        app.setStyle('Fusion')
        font = QFont()
        font.setPointSize(10)
        app.setFont(font)
        window = TestStationInterface()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logger.critical(f"Application crashed: {str(e)}", exc_info=True)
        QMessageBox.critical(None, "Fatal Error", f"The application encountered a fatal error:\n{str(e)}")
        sys.exit(1)
