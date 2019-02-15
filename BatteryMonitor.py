"""
Name: Phi Phan
Date: 10/12/2018
Description:
    The purpose of this script is to monitor and log the voltage and current of the
    batteries used in the products of JSF Technologies. The script will also measure and log values from a thermistor on
    an Arduino The script will provide a GUI for the user to reference the
    measurements as well as log this data on to an .xlsx file. Data points will be logged every 30 minutes.
    The script utilizes modules from Measurement Instruments UL Python Library and Tkinter
    This script is to be used with a Measurement Instruments USB-1208LS.

Notes:
- 12/12/2018 An extra LED Flash command is used in between ul.a_input_mode(self.board_num, AnalogInputMode.SINGLE_ENDED)
 and self.ai_props = AnalogInputProps(self.board_num) to give the board time between configurations.
 This seems to remedy the issue where input channels are swapped.
- 12/18/2018 Adding temperature probe via Arduino. Using Serial to obtain temperature.
            CONFIRM PORT NUMBER AND BAUD RATE OF ARDUINO. UNLESS THIS HAS BEEN CHANGED, THE CORRECT VALUES ARE BELOW
"""

import xlsxwriter
import time
import datetime
import serial
#from serial.serialjava import Serial
from builtins import *
from mcculw import ul
from mcculw.enums import InterfaceType
from mcculw.enums import ULRange
from mcculw.enums import AnalogInputMode
from uiexample import UIExample
from mcculw.examples.props.ai import AnalogInputProps
from mcculw.ul import ULError
import tkinter as tk
from tkinter import messagebox

# Defines
SERIAL_PORT = 'COM95' #CONFIRM COM PORT
# set this to the same rate used on the Arduino
SERIAL_RATE = 115200

# Global Variables
flag = 30 # log time in seconds
row = 5
col = 0
sheet = False


    #Initialize Device
class BatteryMonitor(UIExample):
    def __init__(self,master):
        super(BatteryMonitor, self).__init__(master)
        self.board_num = 0
        ul.ignore_instacal()
        if self.discover_devices():
            self.create_widgets()
        else:
            self.create_unsupported_widgets(self.board_num)

    #Detect Device
    def discover_devices(self):
        # Get the device inventory
        devices = ul.get_daq_device_inventory(InterfaceType.USB)

        #check for USB-1208LS
        self.device = next((device for device in devices
                   if "USB-1208LS" in device.product_name), None)

        if self.device != None:
            # Create the DAQ device from the descriptor
            # For performance reasons, it is not recommended to create and release
            # the device every time hardware communication is required.
            # Create the device once and do not release it
            # until no additional library calls will b e made for this device
            # (typically at application exit).

            ul.create_daq_device(self.board_num, self.device)
            ul.flash_led(self.board_num)
            ul.a_input_mode(self.board_num, AnalogInputMode.SINGLE_ENDED)
            ul.flash_led(self.board_num)
            self.ai_props = AnalogInputProps(self.board_num)
            return True
        return False

    def update_value(self):
        try:
            # Read voltage value from each channel
            global flag
            global row
            global col
            channel = 0
            ai_range = ULRange.BIP10VOLTS
            ser = serial.Serial(SERIAL_PORT, SERIAL_RATE)

            ts = time.time()
            st = datetime.datetime.fromtimestamp(ts).strftime('%H:%M:%S')

            value1 = ul.v_in(self.board_num, channel, ai_range)
            value1 = value1 * 2.06 # Number calculated to compensate for voltage divider.

            value2 = ul.v_in(self.board_num, channel + 1, ai_range)
            value2 = value2 / -9.99 #step down voltage from op-amp input
           # value2 = value2 / 1.1 # Calculate current from voltage value

            value3 = ul.v_in(self.board_num, channel + 2, ai_range)
            value3 = value3 * 2.06  # to remove voltage divider

            value4 = ul.v_in(self.board_num, channel + 3, ai_range)
            value4 = value4 / -10.04 #step down voltage from op-amp input
            #value4 = value4 / 1.1 # Calculate current from voltage value

            value5 = ul.v_in(self.board_num, channel + 4, ai_range)
            value5 = value5 * 2.06  # to remove voltage divider

            value6 = ul.v_in(self.board_num, channel + 5, ai_range)
            value6 = value6 / -10.07 #step down voltage from op-amp input
            #value6 = value6 / 1.1 # Calculate current from voltage value

            value7 = ul.v_in(self.board_num, channel + 6, ai_range)
            value7 = value7 * 2.06  # to remove voltage divider

            value8 = ul.v_in(self.board_num, channel + 7, ai_range)
            value8 = value8 / -10.03 #step down voltage from op-amp input
            #value8 = value8 / 1.1 # Calculate current from voltage value

            # obtain temperature value from thermocouple on Arduino
            tempValue = ser.readline().decode('utf-8')

            #Output values on GUI
            self.BV1["text"] = '{:.2f}'.format(value1) + "V"
            self.BC1["text"] = '{:.3f}'.format(value2) + "A"

            self.BV2["text"] = '{:.2f}'.format(value3) + "V"
            self.BC2["text"] = '{:.3f}'.format(value4) + "A"

            self.BV3["text"] = '{:.2f}'.format(value5) + "V"
            self.BC3["text"] = '{:.3f}'.format(value6) + "A"

            self.BV4["text"] = '{:.2f}'.format(value7) + "V"
            self.BC4["text"] = '{:.3f}'.format(value8) + "A"

            self.temperatureLabel["text"] = str(tempValue)


            # # Write values
            # worksheet.write(row, col, value1)
            # worksheet.write(row, col + 1, value2)
            # worksheet.write(row, col + 2, value3)
            # worksheet.write(row, col + 3, value4)
            # worksheet.write(row, col + 4, value5)
            # worksheet.write(row, col + 5, value6)
            # worksheet.write(row, col + 6, value7)
            # worksheet.write(row, col + 7, value8)
            # worksheet.write(row, col + 8, tempValue)
            # row += 1

            # Write values on to spreadsheet
            if flag >= 30: #Log every 30 seconds
                worksheet.write(row, col, value1)
                worksheet.write(row, col + 1, value2)
                worksheet.write(row, col + 2, value3)
                worksheet.write(row, col + 3, value4)
                worksheet.write(row, col + 4, value5)
                worksheet.write(row, col + 5, value6)
                worksheet.write(row, col + 6, value7)
                worksheet.write(row, col + 7, value8)
                worksheet.write(row, col + 8, tempValue)
                worksheet.write(row, col + 9, st)
                row += 1
                flag = 0
            else:
                flag += 1

            if self.running:
                self.after(1000, self.update_value)
        except ULError as e:
            self.show_ul_error(e)

    def stop(self):
        self.running = False
        self.start_button["command"] = self.start
        self.start_button["text"] = "Start"

    def start(self):
        self.running = True
        self.start_button["command"] = self.stop
        self.start_button["text"] = "Stop"
        self.update_value()

    def quit(self):
        workbook.close()
        messagebox.showwarning("SPREADSHEET CREATED", "Check directory & Record Batterys")
        self.master.destroy()

    def create_widgets(self):
        # '''Create the tkinter UI'''
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.X, anchor=tk.NW)

        # Display Device ID
        device_id_left_label = tk.Label(main_frame)
        self.device_id_label = tk.Label(main_frame)
        device_id_left_label["text"] = "Device ID: " + self.device.unique_id
        device_id_left_label.grid(row=0, column=0, sticky=tk.W, padx=3, pady=3)


        value_battery_label = tk.Label(main_frame)
        value_battery_label["text"] = ("Battery")
        value_battery_label.grid(row=1, column=0, sticky=tk.W, padx=3, pady=3)

        channel_one_label = tk.Label(main_frame)
        channel_one_label["text"] = ("1")
        channel_one_label.grid(row=2, column=0, sticky=tk.W, padx=3, pady=3)

        channel_two_label = tk.Label(main_frame)
        channel_two_label["text"] = ("2")
        channel_two_label.grid(row=3, column=0, sticky=tk.W, padx=3, pady=3)

        channel_three_label = tk.Label(main_frame)
        channel_three_label["text"] = ("3")
        channel_three_label.grid(row=4, column=0, sticky=tk.W, padx=3, pady=3)

        channel_four_label = tk.Label(main_frame)
        channel_four_label["text"] = ("4")
        channel_four_label.grid(row=5, column=0, sticky=tk.W, padx=3, pady=3)

        temp_label = tk.Label(main_frame)
        temp_label["text"] = ("Temperature [C]:")
        temp_label.grid(row=7, column=0, sticky=tk.W, padx=3, pady=3)

        value_voltage_label = tk.Label(main_frame)
        value_voltage_label["text"] = ("Voltage")
        value_voltage_label.grid(row=1, column=1, sticky=tk.W, padx=3, pady=3)

        value_current_label = tk.Label(main_frame)
        value_current_label["text"] = ("Current")
        value_current_label.grid(row=1, column=2, sticky=tk.W, padx=3, pady=3)


        #Label placeholders for ADC values
        self.BV1 = tk.Label(main_frame)
        self.BV1.grid(row=2, column=1, sticky=tk.E, padx=3, pady=3)

        self.BC1 = tk.Label(main_frame)
        self.BC1.grid(row=2, column=2, sticky=tk.E, padx=3, pady=3)

        self.BV2 = tk.Label(main_frame)
        self.BV2.grid(row=3, column=1, sticky=tk.E, padx=3, pady=3)

        self.BC2 = tk.Label(main_frame)
        self.BC2.grid(row=3, column=2, sticky=tk.E, padx=3, pady=3)

        self.BV3 = tk.Label(main_frame)
        self.BV3.grid(row=4, column=1, sticky=tk.E, padx=3, pady=3)

        self.BC3 = tk.Label(main_frame)
        self.BC3.grid(row=4, column=2, sticky=tk.E, padx=3, pady=3)

        self.BV4 = tk.Label(main_frame)
        self.BV4.grid(row=5, column=1, sticky=tk.E, padx=3, pady=3)

        self.BC4 = tk.Label(main_frame)
        self.BC4.grid(row=5, column=2, sticky=tk.E, padx=3, pady=3)

        # Placeholder for thermocoupling
        self.temperatureLabel = tk.Label(main_frame)
        self.temperatureLabel.grid(row=8, column=0, sticky=tk.W, padx=3, pady=3)


        button_frame = tk.Frame(self)
        button_frame.pack(fill=tk.X, side=tk.RIGHT, anchor=tk.SE)

        self.start_button = tk.Button(button_frame)
        self.start_button["text"] = "Start"
        self.start_button["command"] = self.start
        self.start_button.grid(row=0, column=0, padx=3, pady=3)

        self.quit_button = tk.Button(button_frame)
        self.quit_button["text"] = "Quit"
        self.quit_button["command"] = self.quit
        self.quit_button.grid(row=0, column=1, padx=3, pady=3)


# Start the program if this module is being run
if __name__ == "__main__":

    # Create Spreadsheet
    ts = time.time()
    st = datetime.datetime.fromtimestamp(ts).strftime('%d%m%Y %Hh%M')
    st_book = datetime.datetime.fromtimestamp(ts).strftime('%d/%m/%Y %H:%M:%S')
    workbook = xlsxwriter.Workbook(st + '.xlsx')
    worksheet = workbook.add_worksheet()
    cell_format = workbook.add_format()
    cell_format.set_center_across()

    worksheet.merge_range('A1:E1', 'This file logs data of battery values', 0)
    worksheet.write(1, 0, st)
    worksheet.merge_range('A4:B4', 'Battery 1', cell_format)
    worksheet.merge_range('C4:D4', 'Battery 2', cell_format)
    worksheet.merge_range('E4:F4', 'Battery 3', cell_format)
    worksheet.merge_range('G4:H4', 'Battery 4', cell_format)
    worksheet.write(4, 8, 'Temperature [C]')

    worksheet.write(4, 0, 'Voltage[V]')
    worksheet.write(4, 1, 'Current[I]')
    worksheet.write(4, 2, 'Voltage[V]')
    worksheet.write(4, 3, 'Current[I]')
    worksheet.write(4, 4, 'Voltage[V]')
    worksheet.write(4, 5, 'Current[I]')
    worksheet.write(4, 6, 'Voltage[V]')
    worksheet.write(4, 7, 'Current[I]')

    BatteryMonitor(master=tk.Tk()).mainloop()
