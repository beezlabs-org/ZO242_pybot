import time
import win32com.client
import pyautogui
from dataclasses import dataclass
from dmsService import Dms
from tulipService import Bot
from tulipService.model.variableModel import VariableModel
import os
from datetime import datetime, timedelta
import psutil

@dataclass
class BotInputSchema:
    language: VariableModel = None
    serverNameSap: VariableModel = None
    client: VariableModel = None
    beekeeperUrl: VariableModel = None
    dmsCred: VariableModel = None
    sapCreds: VariableModel = None
    transactionCode: VariableModel = None
    variant: VariableModel = None
    startDate: VariableModel = None
    endDate: VariableModel = None
    variantDetailName: VariableModel = None
    DistributionName: VariableModel = None
    selectionscreenName: VariableModel = None
    finalscreenName: VariableModel = None
    filename: VariableModel = None
    reportscreenshot : VariableModel = None
    variantscreen_filename: VariableModel = None
    variantScreenshotVariableName: VariableModel = None
    Distributionscreen_filename:VariableModel = None
    DistributionScreenshotVariableName:VariableModel = None
    selectionscreen_filename:VariableModel = None
    selectionScreenVariableName:VariableModel = None
    reportscreen_filename:VariableModel = None
    reportscreenScreenshotVariableName:VariableModel = None





@dataclass
class BotOutputSchema:
    def __init__(self):
        self.ToExecutionTime = VariableModel()
        self.ToExecutionDate = VariableModel()
        self.fromExecutionTime = VariableModel()
        self.fromExecutionDate = VariableModel()
        self.reportfilesignature = VariableModel()

class BotLogic(Bot):
    def __init__(self) -> None:
        super().__init__()
        try:
            self.outputs = BotOutputSchema()
            self.input = self.bot_input.get_proposedBotInputs(BotInputs=BotInputSchema)
            self.sapIdentity = self.bot_input.get_identity(self.input.sapCreds.value)
            self.dmsIdentity = self.bot_input.get_identity(self.input.dmsCred.value)

            # Launch SAP GUI using saplogon.exe
            appName = "saplogon.exe"
            os.startfile(appName)
            time.sleep(10)  # Increased the sleep time to ensure SAP GUI is fully loaded

            # Connect to SAP GUI
            sapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not isinstance(sapGuiAuto, win32com.client.CDispatch):
                raise RuntimeError("Failed to connect to SAP GUI scripting engine")
            application = sapGuiAuto.GetScriptingEngine
            connection = application.OpenConnection(self.input.serverNameSap.value, True)
            time.sleep(10)  # Increased the sleep time to ensure connection is established
            self.sapGui = connection.Children(0)

            self._dms = Dms(beekeeper_url=self.input.beekeeperUrl.value,
                            user_name=self.dmsIdentity.credential.basicAuth.username,
                            password=self.dmsIdentity.credential.basicAuth.password,
                            logger=self.log)
            self.log.info(message="Connected to DMS")
        except Exception as error:
            self.log.info(f"Error initializing BotLogic: {error}")
            raise

    def get_system_time(self):
        # Get the current system time
        now = datetime.now()
        # Format the time as a string
        current_time = now.strftime("%H:%M:%S")
        return current_time

    def get_current_date(self):
        # Get the current system date
        today = datetime.now()
        # Format the date as a string
        current_date = today.strftime("%Y-%m-%d")
        return current_date

    def get_To_execution_time(self):
        # Get the current system time and add 2 hours
        new_time = datetime.now() + timedelta(hours=2)
        # Format the new time as a string
        To_execution_time = new_time.strftime("%H:%M:%S")
        return To_execution_time

    def get_To_execution_date(self):
        # Get the current system date and add 2 hours
        new_date = datetime.now() + timedelta(hours=2)
        # Format the new date as a string
        To_execution_date = new_date.strftime("%Y-%m-%d")
        return To_execution_date

    def login_to_sap(self):
        try:
            self.sapGui.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.sapIdentity.credential.basicAuth.username
            self.sapGui.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.sapIdentity.credential.basicAuth.password
            time.sleep(2)
            self.sapGui.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
            self.sapGui.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 8
            self.sapGui.findById("wnd[0]").sendVKey(0)
            time.sleep(10)  # Increased the sleep time to ensure login is processed
        except Exception as e:
            self.log.info(f"Error during SAP login: {e}")
            raise

    def maximize_window(self):
        self.sapGui.findById("wnd[0]").maximize()

    def handle_multiple_logins(self):
        if self.sapGui.Children.Count > 1:
            self.log.info("Multiple logins detected. Logging out.")
            return False
        return True

    def enter_transaction_code(self):
        self.sapGui.findById("wnd[0]/tbar[0]/okcd").text = self.input.transactionCode.value
        self.sapGui.findById("wnd[0]/tbar[0]/btn[0]").press()
        time.sleep(5)  # Ensure transaction code is processed

    def open_selection_screen(self):
        self.sapGui.findById("wnd[0]/tbar[1]/btn[17]").press()
        time.sleep(5)

    def enter_selection_criteria(self):
        self.sapGui.findById("wnd[1]/usr/txtV-LOW").text = self.input.variant.value
        time.sleep(1)
        self.sapGui.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        file_name = self.input.variantscreen_filename.value
        time.sleep(5)
        self.take_screenshot(file_name)
        variantScreen_fs = self.upload_to_dms(file_name)
        self.bot_output.add_variable(
            key=self.input.variantScreenshotVariableName.value,
            val=variantScreen_fs,
        )
        time.sleep(1)
        self.sapGui.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(5)

    def convert_date_format(self, date_str):
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            return date_obj.strftime('%d.%m.%Y')
        except ValueError as e:
            self.log.error(f"Error in date format conversion: {e}")
            return None

    def set_dates(self):
        Startdate = self.convert_date_format(self.input.startDate.value)
        self.sapGui.findById("wnd[0]/usr/ctxtSO_DATE-LOW").text = Startdate
        EndDate = self.convert_date_format(self.input.endDate.value)
        self.sapGui.findById("wnd[0]/usr/ctxtSO_DATE-HIGH").text = EndDate
        self.sapGui.findById("wnd[0]/usr/ctxtSO_DATE-HIGH").setFocus()

    def capture_screenshots(self):
        self.sapGui.findById("wnd[0]/usr/btn%_SO_VTWEG_%_APP_%-VALU_PUSH").press()
        time.sleep(5)
        Distributionfile_name = self.input.Distributionscreen_filename.value
        time.sleep(5)
        self.take_screenshot(Distributionfile_name)
        DistributionScreen_fs = self.upload_to_dms(Distributionfile_name)
        self.bot_output.add_variable(
            key=self.input.DistributionScreenshotVariableName.value,
            val=DistributionScreen_fs,
        )
        self.sapGui.findById("wnd[1]/tbar[0]/btn[12]").press()
        time.sleep(5)
        selectionscreen_name = self.input.selectionscreen_filename.value
        time.sleep(5)
        self.take_screenshot(selectionscreen_name)
        selectionScreen_fs = self.upload_to_dms(selectionscreen_name)
        self.bot_output.add_variable(
            key=self.input.selectionScreenVariableName.value,
            val=selectionScreen_fs,
        )
        time.sleep(5)

    def final_screen(self):
        self.sapGui.findById("wnd[0]/mbar/menu[0]/menu[2]").select()
        time.sleep(2)
        self.sapGui.findById("wnd[1]/tbar[0]/btn[13]").press()
        time.sleep(2)
        self.sapGui.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
        time.sleep(2)
        self.sapGui.findById("wnd[1]/tbar[0]/btn[11]").press()
        time.sleep(2)
        self.sapGui.findById("wnd[0]/mbar/menu[2]/menu[0]/menu[1]").select()
        self.sapGui.findById("wnd[1]/usr/txtV-LOW").text = self.input.variant.value
        self.sapGui.findById("wnd[1]/usr/txtV-LOW").caretPosition = 14
        self.sapGui.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(2)
        reportscreenfile_name = self.input.reportscreen_filename.value
        time.sleep(5)
        self.take_screenshot(reportscreenfile_name)
        DistributionScreen_fs = self.upload_to_dms(reportscreenfile_name)
        self.bot_output.add_variable(
            key=self.input.reportscreenScreenshotVariableName.value,
            val=DistributionScreen_fs,
        )
        container_id = "wnd[0]/usr/cntlALV_CONTAINER_3/shellcont/shell"
        menu_item_id = "&XXL"
        control = self.sapGui.findById(container_id)
        control.contextMenu()
        control.selectContextMenuItem(menu_item_id)
        self.sapGui.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.sapGui.findById("wnd[1]/usr/ctxtDY_PATH").text = self.working_dir
        self.sapGui.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.input.filename.value
        self.sapGui.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        self.sapGui.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.outputs.reportfilesignature.value = self._dms.upload_file_to_dms(self.input.filename.value)
        self.log.info(message="This report file has been uploaded to DMS")
        self.close_excel()


    def close_excel(self):
        try:
            for process in psutil.process_iter(['pid', 'name']):
                if process.info['name'] == 'EXCEL.EXE':
                    process.terminate()
                    process.wait(timeout=5)
            self.log.info(message="Excel application closed")
        except Exception as e:
            self.log.error(f"Error closing Excel application: {e}")
            raise


    def log_off_sap(self):
        try:
            self.sapGui.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            self.sapGui.findById("wnd[0]").sendVKey(0)
            self.log.info(message="Logged off from SAP")
        except Exception as e:
            self.log.error(f"Error logging off from SAP: {e}")
            self.bot_output.error()
            raise

    def take_screenshot(self, filename):
        try:
            time.sleep(2)
            fullpath = os.path.join(self.working_dir, filename)
            # Take screenshot
            self.log.info(f"Trying to take screenshot {fullpath}.")
            pyautogui.screenshot(fullpath)
            self.log.info(
                f"Screenshot has been successfully taken {fullpath} and saved."
            )
            return fullpath
        except Exception as e:
            self.log.error(f"Error in line take screenshot: {e}")
            raise

    def upload_to_dms(self, filename):
        try:
            time.sleep(2)
            # Construct the full path of the file
            fullpath = os.path.join(self.working_dir, filename)
            time.sleep(2)
            self.log.info(f"Trying to upload the {fullpath} to dms.")
            # Upload to DMS and returns file signature
            output_fs = self._dms.upload_file_to_dms(fullpath)
            self.log.info(f"Successfully uploaded this {fullpath} to dms.")
            return output_fs
        except Exception as e:
            self.log.error(f"Error in uploading to dms: {e}")
            raise


    def main(self):
        try:
            self.login_to_sap()
            self.maximize_window()
            if not self.handle_multiple_logins():
                return
            self.enter_transaction_code()
            self.open_selection_screen()
            self.enter_selection_criteria()
            self.set_dates()
            self.capture_screenshots()
            self.final_screen()


            # Get ToExecutionTime and ToExecutionDate
            To_execution_time = self.get_To_execution_time()
            To_execution_date = self.get_To_execution_date()
            self.outputs.ToExecutionTime.value = To_execution_time
            self.outputs.ToExecutionDate.value = To_execution_date

            # Calculate fromExecutionTime and fromExecutionDate
            To_execution_datetime = datetime.strptime(f"{To_execution_date} {To_execution_time}", "%Y-%m-%d %H:%M:%S")
            from_execution_datetime = To_execution_datetime - timedelta(minutes=2)
            from_execution_time = from_execution_datetime.strftime("%H:%M:%S")
            from_execution_date = from_execution_datetime.strftime("%Y-%m-%d")

            self.outputs.fromExecutionTime.value = from_execution_time
            self.outputs.fromExecutionDate.value = from_execution_date

            self.log.info(f"To Execution Time (after adding 2 hours): {To_execution_time}")
            self.log.info(f"To Execution Date (after adding 2 hours): {To_execution_date}")
            self.log.info(f"From Execution Time (2 minutes earlier): {from_execution_time}")
            self.log.info(f"From Execution Date (2 minutes earlier): {from_execution_date}")
            self.bot_output.success(message="The z0242 bot completed its work successfully!")

        except Exception as e:
            self.bot_output.error("bot failed")
            self.log.error(f"Error in main logic: {e}")
            raise
        finally:
            self.log_off_sap()

