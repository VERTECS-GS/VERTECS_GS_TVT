import sys
import os
from datetime import datetime
import pandas as pd
from PyQt6 import QtWidgets, uic
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QMessageBox, QFileDialog, QTableWidgetItem
import serial
import serial.tools.list_ports
from threading import Thread
import time
import threading
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import subprocess  # For launching external software
from Track1 import Ui_Track1



class ComPortApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        # uic.loadUi("Track1.ui", self)  # Load the UI
       
        self.ui = Ui_Track1()  # Initialize the UI
        self.ui.setupUi(self)      # Set up the UI on the main window
        self.populate_com_ports()  # Populate the combo box with available COM ports

        # Add a status label to display connection status
        self.status_label = QLabel("Status: Not Connected", self)
        self.status_label.setGeometry(50, 500, 300, 30)  # Adjust as needed
        self.status_label.setStyleSheet("color: white; font-size: 14px;")
        self.status_label.show()
        self.ui.pushButton_3.clicked.connect(self.display_text)

          # Connect the load_s_gse button to the function for launching the software
        self.ui.load_s_gse.clicked.connect(self.launch_software)
    
        
        # Refresh button to reload COM ports
        if hasattr(self, 'refreshButton'):
            self.ui.refreshButton.clicked.connect(self.populate_com_ports)

        # Connect button to connect the selected COM port
        self.ui.connectbutton.clicked.connect(self.connect_to_comport)

        # Send button to send data from commandBox to the COM port (sbandsend)
        if hasattr(self.ui, 'sbandsend'):  # Ensure the UI has a 'sbandsend' button
            self.ui.sbandsend.clicked.connect(self.send_data_to_comport)

        # Connect button for receiving COM port
        self.ui.connectbutton_2.clicked.connect(self.connect_to_receive_port)

        # Refresh baud rate
        if hasattr(self, 'refreshButton_5'):
            self.ui.refreshButton_5.clicked.connect(self.populate_baudrate)

        #texxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxt file save 
        
        # Initialize filename with the start time of the software
        if not hasattr(self, 'log_filename'):
            start_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            self.log_filename = f"serial_log_{start_time}.txt"
            with open(self.log_filename, 'w') as log_file:
                log_file.write(f"Log started at {start_time}\n")

        # Load command button to load Excel data
        self.ui.loadcommand.clicked.connect(self.load_excel_data)

        # Handle table widget item selection
        self.ui.tableWidget.cellClicked.connect(self.display_selected_data)

        # Save button to initialize the save file for sent data
        self.ui.saveButton.clicked.connect(self.initialize_excel_file)



        # Variables for saving data
        self.excel_file_path = None
        self.data_to_save = []  # Temporary buffer to store sent data

        # New variables for additional features
        self.data_buffer = []  # Buffer to hold received data
        self.ui.receivdata.setPlainText("")  # Clear initial QTextEdit content
        
        self.graph_viewer = self.GraphViewer(self)
        self.layout().addWidget(self.graph_viewer)
        self.data_buffer = []
        # self.graph_viewer = GraphViewer(self)
        self.init_ui()
    
    class GraphViewer(FigureCanvas):
            def __init__(self, parent=None):
                self.figure = Figure()
                super().__init__(self.figure)
                self.axes = self.figure.add_subplot(111)
                self.setParent(parent)
                self.setMinimumSize(400, 300)

            def update_graph(self, data):
                """
                Update the graph with new data.
                """
                self.axes.clear()
                self.axes.plot(data, label="Voltage", color='#BC5090')  # Set line color
                self.axes.set_title("RTHK", color='red')  # Set title color
                self.axes.set_xlabel("Pass", color='green')  # Set x-axis label color
                self.axes.set_ylabel("Value", color='purple')  # Set y-axis label color
                self.axes.tick_params(axis='x', colors='green')  # Set x-axis tick color
                self.axes.tick_params(axis='y', colors='purple')  # Set y-axis tick color
                self.axes.legend()
                self.draw()
    
    def init_ui(self):
        # Set the geometry of the graph viewer
        self.graph_viewer.setGeometry (820, 60, 600, 380)  # Example position and size
        # Alternatively, you can use a layout to manage placement
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self.graph_viewer)
        layout.addStretch()  # Add space for better positioning
    
    #loading sband_gse software
    def launch_software(self):
        
        software_path = "s_gse_v1_0.exe"  #software path
        if not os.path.exists(software_path):
            QMessageBox.critical(self, "Error", f"Software not found at {software_path}")
            return

        try:
            subprocess.Popen(software_path, shell=True)  # Launch the software
            QMessageBox.information(self, "Success", f"Software launched: {software_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to launch the software: {e}")


        # Connect the load_s_gse button to the function for launching the software
        self.ui.load_s_gse.clicked.connect(self.launch_software)

   
    def initialize_excel_file(self):
        """
        Initialize the Excel file to save sent data. If a file already exists for today, create a new one with a unique name.
        """
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")  # Format: YYMMDD_HHMMSS
        base_name = f"VERTECS_TRACK_{current_time}"
        counter = 1
        file_name = f"{base_name}.xlsx"

        #if the file already exists
        while os.path.exists(file_name):
            file_name = f"{base_name} ({counter}).xlsx"
            counter += 1
        self.excel_file_path = file_name
        print(f"DEBUG: Excel file path set to: {self.excel_file_path}")  # Added this line for debugging
        self.save_data_to_excel(initial=True)
        QMessageBox.information(self, "Success", f"Excel file initialized: {file_name}")

    def save_data_to_excel(self, initial=False):
        """
        Save the data to the Excel file. If `initial` is True, create the file with headers.
        """
        if not self.excel_file_path:
            QMessageBox.warning(self, "Error", "Please initialize the save file first by clicking the save button.") #ensure for saving file
            return

        if initial:
            # Create a new file with headers
            df = pd.DataFrame(columns=["Serial", "Date", "Time", "Sent Data", "Count"])
            df.to_excel(self.excel_file_path, index=False)

        # Save new data if available
        if self.data_to_save:
            existing_data = pd.read_excel(self.excel_file_path)
            new_data = pd.DataFrame(self.data_to_save, columns=["Serial", "Date", "Time", "Sent Data", "Count"])
            updated_data = pd.concat([existing_data, new_data], ignore_index=True)
            # Update serial numbers for all rows
            updated_data["Serial"] = range(1, len(updated_data) + 1)
            updated_data.to_excel(self.excel_file_path, index=False)
            self.data_to_save = []  # Clear the saved data buffer

    def load_excel_data(self):
        """
        Load data from an Excel file and populate the tableWidget.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                self.excel_data = pd.read_excel(file_path)  # Load Excel data into a DataFrame
                self.populate_table(self.excel_data)  # Populate the tableWidget with data
                QMessageBox.information(self, "Success", "Excel file loaded successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load Excel file: {e}")

    def populate_table(self, data):
        """
        Populate the tableWidget with the given DataFrame.
        """
        self.ui.tableWidget.setRowCount(0)  # Clear existing rows
        self.ui.tableWidget.setColumnCount(len(data.columns))  # Set number of columns
        self.ui.tableWidget.setHorizontalHeaderLabels(data.columns)  # Set column headers

        for row_index, row in data.iterrows():
            self.ui.tableWidget.insertRow(row_index)  # Add a new row
            for col_index, value in enumerate(row):
                self.ui.tableWidget.setItem(row_index, col_index, QTableWidgetItem(str(value)))

    def display_selected_data(self, row, column):
        """
        Display the selected cell's data in the Sbo QLineEdit.
        """
        selected_data = self.ui.tableWidget.item(row, column).text()
        self.ui.Sbo.setText(selected_data)

    def send_data_to_comport(self):
        """
        Send data from Sbo to the selected COM port and save it to the Excel file.
        """
        if not hasattr(self, 'serial_connection') or not self.serial_connection.is_open:
            QMessageBox.warning(self, "Connection Error", "No COM port connected.")
            return

        data_to_send = self.ui.Sbo.text().strip()  # Get the text from Sbo
        if not data_to_send:
            QMessageBox.warning(self, "Input Error", "Please enter a command to send.")
            return

        try:
            # Send data to the COM port
            self.serial_connection.write(data_to_send.encode('utf-8'))
            QMessageBox.information(self, "Success", f"Data sent: {data_to_send}")

            # Save data to the internal buffer
            current_date = datetime.now().strftime("%Y-%m-%d")
            current_time = datetime.now().strftime("%H:%M:%S")
# new method for saving data

        # Load existing data from the Excel file
            if os.path.exists(self.excel_file_path):
                existing_data = pd.read_excel(self.excel_file_path)

            # Check if the data already exists in the file
                matching_rows = existing_data[existing_data["Sent Data"] == data_to_send]
                

                if not matching_rows.empty:
                    count = matching_rows["Count"].max() + 1
                else:
                    count= 1
                    # Update the count for the matching row
                    
            
                # Add the new row with the updated count
                new_row = {
                    "Serial": len(existing_data) + 1,
                    "Date": current_date,
                    "Time": current_time,
                    "Sent Data": data_to_send,
                    "Count": count,
                }
                updated_data = pd.concat([existing_data, pd.DataFrame([new_row])], ignore_index=True)
            else:
                # If the file does not exist, create it with the initial row
                updated_data = pd.DataFrame(
                    [{
                        "Serial": 1,
                        "Date": current_date,
                        "Time": current_time,
                        "Sent Data": data_to_send,
                        "Count": 1,
                    }]
                )
                self.excel_file_path = f"VERTECS_TRACK_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"  
                  
            # Save the updated data to the Excel file
            updated_data.to_excel(self.excel_file_path, index=False)
            #QMessageBox.information(self, "Success", "Data saved successfully.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to send data: {e}")


    

    def display_text(self):
        # Define the text to display
        text_to_display = "FA F3 20 56 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 E7 49"
        # Set the text in the Sbo QLineEdit
        self.ui.Sbo.setText(text_to_display)

    def connect_to_comport(self):
        """
        Connect to the COM port specified in the comboBox widget.
        """
        com_port = self.ui.comboBox.currentText()  # Get the selected COM port from the comboBox
        baudrate = self.ui.comboBox_4.currentText()

        if not com_port:
            QMessageBox.warning(self, "Input Error", "Please select a valid COM port.")
            return

        try:
            # Attempt to connect to the COM port
            self.serial_connection = serial.Serial(com_port, baudrate=int(baudrate), timeout=1)
            self.status_label.setText(f"Status: Connected to {com_port}")
            QMessageBox.information(self, "Success", f"Successfully connected to {com_port} with {baudrate} baudrate")
        except serial.SerialException as e:
            self.status_label.setText("Status: Connection Failed")
            QMessageBox.critical(self, "Connection Error", f"Failed to connect to {com_port}\n{e}")

    def connect_to_receive_port(self):
        """
        Connect to the receiving COM port and start listening for data.
        """
        receive_port = self.ui.comboBox_5.currentText()
        baudrate = self.ui.comboBox_6.currentText()

        if not receive_port:
            QMessageBox.warning(self, "Input Error", "Please select a valid COM port for receiving.")
            return

        try:
            # Connect to the receiving COM port
            self.receive_connection = serial.Serial(receive_port, baudrate=int(baudrate), timeout=1)
            self.status_label.setText(f"Status: Connected to Receive Port {receive_port}")
            QMessageBox.information(self, "Success", f"Connected to {receive_port} with {baudrate} baudrate for receiving.")

            # Start a thread to receive data instantly
            self.receive_thread = Thread(target=self.receive_data, daemon=True)
            self.receive_thread.start()
        except serial.SerialException as e:
            self.status_label.setText("Status: Receive Port Connection Failed")
            QMessageBox.critical(self, "Connection Error", f"Failed to connect to {receive_port}\n{e}")

    # def receive_data(self):
    # # """
    # # Continuously receive data from the connected port and display it in the QTextEdit.
    # # """
    #     while self.receive_connection.is_open:
    #         try:
    #             # Read up to 200 bytes from the serial port
    #             data = self.receive_connection.read(200)
    #             if data:
    #                 # Split the data into individual bytes for clarity
    #                 byte_list = list(data)
                    
    #                 # Print the entire array contents in hex format
    #                 print("Current Array in Hex:")
    #                 print([f"{byte:04X}" for byte in byte_list])
                   
    #                 # Remove spaces from the received bytes
    #                 byte_list = [byte for byte in byte_list if byte != 32]  # 32 is the ASCII code for space

    #                 # Combine every two bytes into a single byte-like integer
    #                 byte_list = [byte_list[i] << 8 | byte_list[i + 1] for i in range(0, len(byte_list) - 1, 2)]

    #                 # Append each byte to the buffer as a string representation
    #                 self.data_buffer.extend([str(byte) for byte in byte_list])

                   
                
    #                 # Print the entire array contents in both decimal and hex formats
    #                 print("Current Array State:")
    #                 for i, byte in enumerate(byte_list):
    #                     print(f"Index {i}: {byte} (Hex: {byte:04X})")

    #                 # Display the received data in QTextEdit as decoded text
    #                 if isinstance(data, bytes):
    #                     decoded_data = data.decode('utf-8', errors='ignore')
    #                 self.receivdata.append(decoded_data)

    #                 # Check the 6th index for the specific value and execute LED ON action
    #                 if len(self.data_buffer) > 5 and self.data_buffer[5] == "68":
    #                     print("LED ON")  # Perform the LED ON action

    #                 # Clear the buffer if it reaches the size of 200
    #                 if len(self.data_buffer) >= 200:
    #                     self.data_buffer = []

    #         except Exception as e:
    #             self.status_label.setText(f"Receive Error: {e}")
    #             break

    #         time.sleep(0.1)
###############################################--------------------this part work perfectly ----------------#########################################
    # def receive_data(self):
    #     # """
    #     # Continuously receive ASCII data from the connected port, convert every two ASCII bytes into a single hex value,
    #     # store them in an array, and display in QTextEdit.
    #     # """
    #     while self.receive_connection.is_open:
    #         try:
    #             # Read up to 200 bytes from the serial port as ASCII data
    #             data = self.receive_connection.read(1024)
    #             if data:
    #                 # Decode the ASCII data to a string
    #                 if isinstance(data, bytes):
    #                     ascii_data = data.decode('ascii', errors='ignore')

    #                 # Split ASCII data into characters
    #                 ascii_chars = list(ascii_data)

    #                 # Remove spaces from the ASCII characters
    #                 ascii_chars = [char for char in ascii_chars if char != ' ']

    #                 # Convert every two ASCII characters to a single hex value
    #                 hex_values = []
    #                 for i in range(0, len(ascii_chars) - 1, 2):
    #                     combined_ascii = ascii_chars[i] + ascii_chars[i + 1]
    #                     hex_value = int(combined_ascii.encode('ascii').hex(), 16)
    #                     hex_values.append(hex_value)

    #                 # Append the hex values to the data buffer
    #                 self.data_buffer.extend(hex_values)

    #                 # Print the entire array contents in ASCII
    #                 print("Current Array in ASCII:")
    #                 print(ascii_chars)

    #                 # Print each ASCII character individually
    #                 print("Individual ASCII Characters:")
    #                 for i, char in enumerate(ascii_chars):
    #                     print(f"Index {i}: {char}")

    #                 # Print the entire array contents in hex format
    #                 print("Current Array in Hex:")
    #                 print([f"{hex_value:02X}" for hex_value in hex_values])

    #                 # Print each hex value individually
    #                 print("Individual Hex Values:")
    #                 for i, hex_value in enumerate(hex_values):
    #                     print(f"Index {i}: {hex_value:02X}")

    #                 # Print concatenated hex values as ASCII if hex_values is not empty
    #                 if hex_values:
    #                     concatenated_hex = ''.join(f"{hex_value:02X}" for hex_value in hex_values)
    #                     print("Concatenated Hex as ASCII:", bytes.fromhex(concatenated_hex).decode('ascii', errors='ignore'))

    #                 # Display the received data in QTextEdit as ASCII string
    #                 self.receivdata.append(ascii_data)

    #                 # Check the 6th index for the specific value and execute LED ON action
    #                 if len(hex_values) > 5 and hex_values[6] == 0x4142:  # 0x68 is 'h' in hex
    #                     print("LED ON")  # Perform the LED ON action

    #                 # Clear the buffer if it reaches the size of 200
    #                 if len(self.data_buffer) >= 200:
    #                     self.data_buffer = []

    #         except Exception as e:
    #             self.status_label.setText(f"Receive Error: {e}")
    #             print(f"Receive Error: {e}") 
    #             break

    #         time.sleep(0.1)
###############################################--------------------this part work perfectly end----------------#########################################
    def receive_data(self):
        """
        Continuously receive ASCII data from the connected port, convert every two ASCII bytes into a single hex value,
        store them in an array, and display in QTextEdit.
        """
        while self.receive_connection.is_open:
            try:
                # Read up to 1024 bytes from the serial port as ASCII data
                data = self.receive_connection.read(1024)
                if data:
                    # Decode the ASCII data to a string
                    if isinstance(data, bytes):
                        ascii_data = data.decode('ascii', errors='ignore')

                    # Split ASCII data into characters
                    ascii_chars = list(ascii_data)

                    # Remove spaces from the ASCII characters
                    ascii_chars = [char for char in ascii_chars if char != ' ']
                    # Concatenate every two ASCII characters
                    concatenated_ascii = []
                    for i in range(0, len(ascii_chars) - 1, 2):
                        concatenated_pair = ascii_chars[i] + ascii_chars[i + 1]
                        concatenated_ascii.append(concatenated_pair)

                    # Print the concatenated ASCII pairs
                    print("Concatenated ASCII Pairs:")
                    print(concatenated_ascii)
                    
                    # Convert every two ASCII characters to a single hex value
                    # Convert every two ASCII characters to a single hex value
                    hex_values = []
                    for pair in concatenated_ascii:
                        # Directly get the integer representation of the concatenated ASCII pair
                        hex_value = int(pair, 16)
                        hex_values.append(hex_value)
                        # Append the hex values to the data buffer
                        self.data_buffer.extend(hex_values)

                    # Print the entire array contents in ASCII
                    # print("Current Array in ASCII:")
                    # print(ascii_chars)

                    # Print each ASCII character individually
                    # print("Individual ASCII Characters:")
                    # for i, char in enumerate(ascii_chars):
                    #     print(f"Index {i}: {char}")

                    # Print the entire array contents in hex format
                    print("Current Array in Hex:")
                    print([f"{hex_value:02X}" for hex_value in hex_values])

                    # Print each hex value individually
                    print("Individual Hex Values:")
                    for i, hex_value in enumerate(hex_values):
                        print(f"Index {i}: {hex_value:02X}")

                    # # Print concatenated hex values as ASCII if hex_values is not empty
                    # if hex_values:
                    #     concatenated_hex = ''.join(f"{hex_value:02X}" for hex_value in hex_values)
                    #     print("Concatenated Hex as ASCII:", bytes.fromhex(concatenated_hex).decode('ascii', errors='ignore'))

                    # Display the received data in QTextEdit as ASCII string
                    self.ui.receivdata.append(ascii_data)
                    print(f"10th Value: {hex_values[6]:02X}")
                    # Check the 6th index for the specific value and execute LED ON action
                    if len(hex_values) > 5 and hex_values[6] == 0xAB:  # 0x4142 is "AB" in hex
                        print("LED ON")  # Perform the LED ON action
                        # Start a new thread to process hex values
                        processing_thread = threading.Thread(target=self.process_hex_values, args=(hex_values,))
                        processing_thread.start()

                    # Clear the buffer if it reaches the size of 200
                    if len(self.data_buffer) >= 200:
                        self.data_buffer = []
                        
                    # Save the received ASCII data to the log file with a timestamp
                    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    with open(self.log_filename, 'a') as log_file:
                        log_file.write(f"[{current_time}] {ascii_data}\n")
                    

            except Exception as e:
                self.status_label.setText(f"Receive Error: {e}")
                print(f"Receive Error: {e}")
                break

            time.sleep(0.5)

    def process_hex_values(self, hex_values):
        """
        Process hex values: display data in the text window based on the 10th value
        and update the graph viewer based on the 12th value.
        Split the hex values into two parts.
        """
        try:
            # Initialize or update the graph data list
            if not hasattr(self, 'graph_data_list'):
                self.graph_data_list = []
            # Display data in the text window based on the 10th value
            # if len(hex_values) > 9:
            #     # self.text_window.append(f"10th Value: {hex_values[9]:02X}")
            #     self.Sbo.setText(f"10th Value: {hex_values[9]:02X}")

            # # Update the graph viewer based on the 12th value
            if len(hex_values) > 11:
                self.graph_data_list.append(hex_values[11])  # Append new data
                # self.Sbo.setText(f"10th Value: {hex_values[11]}")
                self.ui.vd_1.setText(f"{hex_values[11]} V")
                self.graph_viewer.update_graph(self.graph_data_list)  # Update graph with all data
                # graph_data = [hex_values[11]]
                # self.graph_viewer.update_graph(graph_data)
            # Split hex values into two parts
            middle_index = len(hex_values) // 2
            first_part = hex_values[:middle_index]
            second_part = hex_values[middle_index:]

            print("First Part of Hex Values:", [f"{value:02X}" for value in first_part])
            print("Second Part of Hex Values:", [f"{value:02X}" for value in second_part])

        except Exception as e:
            print(f"Error in processing hex values: {e}")

    def save_qtextedit_to_excel(self):
        """
        Save the content of the QTextEdit (receivdata) to an Excel file.
        """
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Data to Excel", "", "Excel Files (*.xlsx)")
        if file_path:
            try:
                data = self.ui.receivdata.toPlainText().splitlines()  # Extract text data from QTextEdit
                df = pd.DataFrame(data, columns=["Received Data"])
                df.to_excel(file_path, index=False)
                QMessageBox.information(self, "Success", "Data saved to Excel successfully.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save data: {e}")

    def closeEvent(self, event):
        """
        Ensure the serial connection is closed when the application exits.
        """
        if hasattr(self, 'serial_connection') and self.serial_connection.is_open:
            self.serial_connection.close()
        if hasattr(self, 'receive_connection') and self.receive_connection.is_open:
            self.receive_connection.close()
        if self.data_to_save:
            self.save_data_to_excel()  # Save any remaining data
        event.accept()

    def populate_com_ports(self):
        """
        Populate the comboBox with the available COM ports.
        """
        self.ui.comboBox.clear()  # Clear current items in the combo box
        self.ui.comboBox_5.clear()  # Clear receiving COM port items
        ports = serial.tools.list_ports.comports()
        for port in ports:
            self.ui.comboBox.addItem(port.device)  # Add each port to the main comboBox
            self.ui.comboBox_5.addItem(port.device)  # Add each port to the receiving comboBox
        if not ports:
            QMessageBox.information(self, "No Ports Found", "No COM ports detected.")

    def populate_baudrate(self):
        """
        Populate the baudrate options.
        """
        self.ui.comboBox_4.addItems(["115200", "9600"])  # Main port baudrate options
        self.ui.comboBox_6.addItems(["115200", "9600"])  # Receiving port baudrate options


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    main_window = ComPortApp()
    main_window.show()
    sys.exit(app.exec())
