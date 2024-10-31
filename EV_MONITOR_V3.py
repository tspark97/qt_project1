from PyQt5 import QtCore
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets, uic
from PyQt5 import uic
import sys
import datetime
import cv2
import numpy as np
from pymodbus.client import ModbusTcpClient
from pymodbus.transaction import *
import pandas as pd
import datetime
import cv2
import struct
import pyodbc
import os
import shutil
import requests
import threading
import time
driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
# ================ new parameter 2024.10.05
# 1. Add Barcode reading
# 2. Add Tack Time
# 3. 


class USER_window(QMainWindow):
    def __init__(self):
        super(USER_window, self).__init__()
        uic.loadUi('./GUI/MONITOR_GUI_V2.ui', self)
        self.show()


        #ip, port address, variable # change_line
        self.ip = ''
        self.device_operation = 0
        self.device_error = 0
        self.device_st = 0
        self.device_barcode = 0
        self.device_plasma = 0
        self.location = []
        self.port = 0
        self.url = ""
        self.linename = "
        self.ui_start = 0

        #accdb file create # change_line
        self.template_file_path = 0

        # accdb file create / error
        self.db_location_error = 0
        self.filename = 0
        self.db_file_path = 0
        self.cnxn = None

        # accdb file create / tacktime
        self.db_location_st = 0
        self.cnxn_st = None
        # accdb file create / barcode
        self.db_location_barcode = 0
        self.cnxn_barcode = None
        # accdb file create / plasma welding
        self.db_location_plasma = 0
        self.cnxn_plasma = None
        self.parameter_set() # set parameter : ip, device, location, template_file_fath, db_location

        #accdb file add line
        self.NUM = 0
        self.LOCATION = 0
        self.EVENT = 0
        self.CODE = 0
        self.OCC_TIME = 0

        #sand url data
        self.sand_data_list = []

        #data parameter
        self.pre_event_type = []
        self.event_type = []
        self.event_type_for_label = []

        self.pre_err_code = []
        self.err_code = []

        self.last_event_time = []

        #2024-10-24 barcode, tack time 추가
        self.pre_tack_time = []
        self.tack_time = []

        self.OCC_TIME_ST = 0
        self.LOCATION_ST = 0
        self.TACK_TIME = 0


        self.pre_barcode = [0]
        self.barcode = [1]

        self.OCC_TIME_BARCODE = 0
        self.BARCODE = 0

        self.pre_plasma = [0, 0, 0]
        self.plasma = [0, 0, 0]

        self.pre_model = [0]
        self.model = [0]

        self.OCC_TIME_PLASMA = 0
        self.AMPERE = 0
        self.WELDING_TIME = 0
        self.GAS = 0
        self.MODEL_TYPE = 0

        for err_code_setting in range(len(self.location)):
            self.pre_event_type.append(-1)
            self.event_type.append(0)
            self.err_code.append(0)
            self.last_event_time.append(1)
            self.event_type_for_label.append(0)
            self.pre_err_code.append(0)

            self.pre_tack_time.append(0)
            self.tack_time.append(1)

        # first label set
        self.location_set()

        # self.errorcode_set()
        self.status_set()
        self.errortime_set()

        self.start_status = 1 # START ONLY ONE

        # sand data thread
        self.timer1 = QTimer(self)
        self.timer1.timeout.connect(self.sand_data_ERROR)
        self.timer1.start(100)  # 100ms마다 draw_ui 호출
        print("START MULTIPROCESSING - 1")

        # QTimer 설정 (100ms마다 draw_ui 호출)
        self.timer2 = QTimer(self)
        self.timer2.timeout.connect(self.draw_ui)
        self.timer2.start(1000)  # 100ms마다 draw_ui 호출
        print("START MULTIPROCESSING - 2")

        self.run()
        print("system ready")

    def plasma_scale(self, plasma):
        ampere = plasma[0]
        welding_time = plasma[1]
        gas_value = plasma[2]

        ampere_input_min = 0
        ampere_input_max = 65535
        ampere_output_min = 3
        ampere_output_max = 220
        ampere = ((ampere - ampere_input_min) / (ampere_input_max - ampere_input_min) * (ampere_output_max - ampere_output_min) + ampere_output_min )
        ampere = str(round(ampere)) + "A"

        welding_time = welding_time * 0.1
        welding_time = str(welding_time) + "sec"
        gas_input_min = 0
        gas_input_max = 65535
        gas_output_min = 2
        gas_output_max = 100
        gas_value = ((gas_value - gas_input_min) / (gas_input_max - gas_input_min) * (
                    gas_output_max - gas_output_min) + gas_output_min)
        gas_value = round(gas_value) * 0.1
        gas_value = str(gas_value) + "L"

        plasma_data = [ampere, welding_time, gas_value]

        return plasma_data

    def draw_ui(self):
        self.errorcode_set()
        self.status_set()
        self.errortime_set()
        self.TT_set()
    def sand_data_ERROR(self):
        if len(self.sand_data_list) > 0:
            self.sand_data_url(self.sand_data_list[0])
            self.sand_data_list.pop(0)


    def sand_data_url(self, data_list):
        #monitor?line=AF7A-1&location=Cylinder Assembly&event=RUN&code=0&time=20241231235959
        #(self.LOCATION, self.ERROR_CODE, self.ERROR_TIME_START, self.ERROR_TIME_CLEAR))

        D1 = "line="+ str(data_list[0]) + "&"
        D2 = "location="+ str(data_list[1]) + "&"
        D3 = "event="+ str(data_list[2]) + "&"
        D4 = "code=" + str(data_list[3]) + "&"
        D5 = "time=" + str(data_list[4])

        data = D1+D2+D3+D4+D5
        # print(data)
        response = requests.post(self.url, data=data)
        print("sand data : " + data)
        print("sand url result : ",end="")
        print(response.status_code)

    def parameter_set(self):
        df = pd.read_excel('./SAMPLE.xlsx', header=None)
        df = df.values.flatten().tolist()
        for i in range(len(df)):
            if df[i] == 'IP':
                self.ip = str(df[i+1])
            elif df[i] == 'DEVICE_OPERATION':
                self.device_operation = int(df[i+1])
            elif df[i] == 'DEVICE_ERROR':
                self.device_error = int(df[i+1])
            elif df[i] == 'LOCATION' and df[i+1] != 0 and df[i+1] != '0':
                self.location.append(df[i+1])
            elif df[i] == 'template':
                self.template_file_path = str(df[i+1])
            elif df[i] == 'db_location_error':
                self.db_location_error = str(df[i+1])
            elif df[i] == 'db_location_st':
                self.db_location_st = str(df[i+1])
            elif df[i] == 'db_location_barcode':
                self.db_location_barcode = str(df[i+1])
            elif df[i] == 'db_location_plasma':
                self.db_location_plasma = str(df[i+1])
            elif df[i] == 'PORT':
                self.port = int(df[i+1])
            elif df[i] == 'LINE':
                self.linename = str(df[i+1])
            elif df[i] == 'DEVICE_ST':
                self.device_st = int(df[i+1])
            elif df[i] == 'DEVICE_BARCODE':
                self.device_barcode = int(df[i+1])
            elif df[i] == 'DEVICE_PLASMA':
                self.device_plasma = int(df[i+1])
    def barcode_set(self, barcode):
        value = barcode
        str_1 = ""
        for i in value:
            if i != 0:
                high_byte = chr((i >> 8) & 0xFF)
                low_byte = chr(i & 0xFF)
                ch_temp = low_byte + high_byte
                str_1 = str_1 + ch_temp
        return str_1
    def db_connection_test(self, cnxn):
        try:
            # 연결이 끊어졌거나 유효하지 않을 경우 새로운 연결 시도
            if cnxn is None or cnxn.closed:
                print("DB 연결이 끊겼거나 유효하지 않습니다. 재연결 시도 중...")
                driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
                conn_str = f"DRIVER={driver};DBQ={self.db_file_path};"
                cnxn = pyodbc.connect(conn_str)
                print("DB에 성공적으로 재연결되었습니다.")
        except pyodbc.Error as e:
            print(f"DB 연결 오류: {e}")
            raise

    def db_file_set_first(self,file_path, order):
        driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
        conn_str = f"DRIVER={driver};DBQ={file_path};"
        if order == 0:
            self.cnxn = pyodbc.connect(conn_str)
        elif order == 1:
            self.cnxn_st = pyodbc.connect(conn_str)
        elif order == 2:
            self.cnxn_barcode = pyodbc.connect(conn_str)
        elif order == 3:
            self.cnxn_plasma = pyodbc.connect(conn_str)

    def add_data_accdb(self, cnxn, order):
        # driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
        # conn_str = f"DRIVER={driver};DBQ={file_path};"
        # cnxn = pyodbc.connect(conn_str)
        print(cnxn)
        try:
            # self.db_connection_test(cnxn)  # 연결 상태 확인 및 필요 시 재연결
            cursor = cnxn.cursor()
            if order == 0:
                cursor.execute("""
                INSERT INTO MONITOR (OCC_TIME, LINE, LOCATION, EVENT, CODE)
                VALUES (?, ?, ?, ?, ?)
                """, (self.OCC_TIME, self.linename, self.LOCATION, self.EVENT, self.CODE))

            elif order == 1:
                self.OCC_TIME_ST = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                INSERT INTO TACKTIME (OCC_TIME, LINE, LOCATION, ST)
                VALUES (?, ?, ?, ?)
                """, (self.OCC_TIME_ST, self.linename, self.LOCATION_ST, self.TACK_TIME))
            elif order == 2:
                self.OCC_TIME_BARCODE = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                INSERT INTO BARCODE (OCC_TIME, LINE, BARCODE)
                VALUES (?, ?, ?)
                """, (self.OCC_TIME_BARCODE, self.linename, self.BARCODE))
            elif order == 3:
                self.OCC_TIME_PLASMA = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute("""
                INSERT INTO PLASMA (OCC_TIME, LINE, AMPERE, WELDING_TIME, GAS)
                VALUES (?, ?, ?, ?, ?)
                """, (self.OCC_TIME_PLASMA, self.linename, self.AMPERE, self.WELDING_TIME, self.GAS))

            cnxn.commit()
            print("데이터가 성공적으로 삽입되었습니다.")
        except pyodbc.Error as e:
            print(f"데이터 삽입 중 오류 발생: {e}")
            self.handle_db_error(e)  # 에러 처리 로직 추가


    def handle_db_error(self, error):
        # 데이터베이스 에러 처리 로직
        print(f"데이터베이스 오류 발생: {error}")
        # 연결이 끊겼을 경우, 재연결을 시도하거나 경고 메시지 출력
        if "Lost connection" in str(error):
            print("연결이 끊겼습니다. 재연결 시도 중...")
            time.sleep(2)  # 잠시 대기 후 재연결 시도
            self.db_connection_test(cnxn)  # 재연결 시도
        else:
            # 심각한 오류라면 프로그램을 중단하거나 추가적인 로직을 처리
            print("심각한 데이터베이스 오류가 발생했습니다. 로그를 확인하십시오.")

    def create_file_name(self, order):
        tm = datetime.datetime.now()
        year = str(tm.year)

        if tm.month < 10:
            month = "0" + str(tm.month)
        else:
            month = str(tm.month)

        if tm.day < 10:
            day = "0" + str(tm.day)
        else:
            day = str(tm.day)

        if tm.hour < 10:
            hour = "0" + str(tm.hour)
        else:
            hour = str(tm.hour)

        if tm.minute < 10:
            minute = "0" + str(tm.minute)
        else:
            minute = str(tm.minute)

        if tm.second < 10:
            second = "0" + str(tm.second)
        else:
            second = str(tm.second)
        # type = ".xlsx"
        type = ".accdb"
        if order == 0:
            filename = year + month + day + hour + minute + second + type
        elif order == 1:
            filename = year + month + day + hour + minute + second
        elif order == 2:
            filename = year + "-" + month + "-" + day + " " + hour + ":" + minute
        # print(filename)
        return filename

    def create_accdb_from_template(self, template_path, new_db_path):
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
        shutil.copyfile(template_path, new_db_path)
        print(f"{new_db_path} 파일이 성공적으로 생성되었습니다.")


    def errortime_set(self):
        length = len(self.location)
        label_name = [f'label_errortime_{i}' for i in range(1, length + 1)]
        for i, name in enumerate(label_name):
            label = self.findChild(QtWidgets.QLabel, name)
            if label:
                time_value = str(self.last_event_time[i])
                label.setText(time_value)
    def status_set(self):
        length = len(self.location)
        label_name = [f'label_status_{i}' for i in range(1, length + 1)]

        for i, name in enumerate(label_name):
            label = self.findChild(QtWidgets.QLabel, name)
            if label:
                if self.event_type_for_label[i] == 1:
                    label.setText("RUN")
                    label.setStyleSheet("background-color: rgb(255, 255, 255); color:rgb(0, 0, 0);")
                elif self.event_type_for_label[i] == 0:
                    label.setText("DOWN")
                    label.setStyleSheet("background-color: rgb(255, 0, 0); color:rgb(255, 255, 255);")
                elif self.event_type_for_label[i] == -1:
                    label.setText("IDLE")
                    label.setStyleSheet("background-color: rgb(0, 0, 255); color:rgb(255, 255, 255);")


    def errorcode_set(self):
        length = len(self.location)
        label_name = [f'label_errorcode_{i}' for i in range(1, length + 1)]

        for i, name in enumerate(label_name):
            label = self.findChild(QtWidgets.QLabel, name)
            if label:
                label.setText(str(self.err_code[i]))

    def TT_set(self):
        length = len(self.location)
        label_name = [f'label_tack_{i}' for i in range(1, length + 1)]

        for i, name in enumerate(label_name):
            label = self.findChild(QtWidgets.QLabel, name)
            if label:
                label.setText(str(self.tack_time[i]))

    def location_set(self):
        length = len(self.location)
        label_name = [f'label_location_{i}' for i in range(1, length+1)]

        for i, name in enumerate(label_name):
            label = self.findChild(QtWidgets.QLabel, name)
            if label:
                label.setText(self.location[i])

    def run(self):
        self.filename = self.create_file_name(order=0)
        location_lst = [self.db_location_error, self.db_location_st, self.db_location_barcode, self.db_location_plasma]
        for g in range(len(location_lst)):
            self.db_file_set_first(file_path=location_lst[g], order=g)


        self.label_program_status.setText("RUN")
        self.start_status = 1
        print('start run')

        while self.start_status:
            cv2.waitKey(100)

            self.err_code, self.event_type, self.tack_time, self.barcode, self.plasma = self.read_resister()

            self.barcode = self.barcode_set(barcode=self.barcode)

            self.plasma = self.plasma_scale(self.plasma)

            if self.pre_plasma != self.plasma:
                self.AMPERE = self.plasma[0]
                self.WELDING_TIME = self.plasma[1]
                self.GAS = self.plasma[2]

                self.add_data_accdb(self.cnxn_plasma, 3)

                self.pre_plasma = self.plasma

                self.label_ampere.setText(self.plasma[0])
                self.label_welding_time.setText(self.plasma[1])
                self.label_gas.setText(self.plasma[2])

            if self.pre_barcode != self.barcode:
                self.BARCODE = str(self.barcode)
                self.add_data_accdb(cnxn=self.cnxn_barcode, order=2)
                self.pre_barcode = self.barcode
                self.label_barcode.setText(str(self.barcode))

            for n in range(len(self.location)):
                # RUN, DOWN, IDLE 처리
                if self.pre_tack_time[n] != self.tack_time[n]:
                    self.LOCATION_ST = self.location[n]
                    self.TACK_TIME = str(self.tack_time[n])
                    self.add_data_accdb(cnxn=self.cnxn_st, order=1)

                if self.pre_event_type[n] != self.event_type[n]:
                    if self.event_type[n] == 1 and self.err_code[n] == 0:
                        print("run")
                        self.last_event_time[n] = self.create_file_name(order=2)
                        self.event_type_for_label[n] = 1
                        self.NUM += 1
                        self.LOCATION = self.location[n]
                        self.EVENT = "run"
                        self.CODE = self.err_code[n]
                        self.OCC_TIME = self.create_file_name(order=1)
                        self.data_list = [self.linename, self.LOCATION, self.EVENT, self.CODE, self.OCC_TIME]
                        self.sand_data_list.append(self.data_list)
                        self.add_data_accdb(cnxn=self.cnxn, order=0)

                    elif self.event_type[n] == 0:
                        print("down")
                        self.last_event_time[n] = self.create_file_name(order=2)
                        self.event_type_for_label[n] = 0
                        self.NUM += 1
                        self.LOCATION = self.location[n]
                        self.EVENT = "down"
                        self.CODE = self.err_code[n]
                        self.OCC_TIME = self.create_file_name(order=1)
                        self.data_list = [self.linename, self.LOCATION, self.EVENT, self.CODE, self.OCC_TIME]
                        self.sand_data_list.append(self.data_list)
                        self.add_data_accdb(cnxn=self.cnxn, order=0)

                elif self.event_type[n] == 1 and self.pre_event_type[n] == 1 and (self.err_code[n] == 99 or self.err_code[n]==98):
                    if self.event_type_for_label[n] != -1:
                        print("idle")
                        self.last_event_time[n] = self.create_file_name(order=2)
                        self.event_type_for_label[n] = -1
                        self.NUM += 1
                        self.LOCATION = self.location[n]
                        self.EVENT = "idle"
                        self.CODE = self.err_code[n]
                        self.OCC_TIME = self.create_file_name(order=1)
                        self.data_list = [self.linename, self.LOCATION, self.EVENT, self.CODE, self.OCC_TIME]
                        self.sand_data_list.append(self.data_list)
                        self.add_data_accdb(cnxn=self.cnxn, order=0)

                elif self.event_type[n] == 1 and self.pre_event_type[n] == 1 and (self.err_code[n] == 0):
                    if self.event_type_for_label[n] != 1:
                        print("run")
                        self.last_event_time[n] = self.create_file_name(order=2)
                        self.event_type_for_label[n] = 1
                        self.NUM += 1
                        self.LOCATION = self.location[n]
                        self.EVENT = "run"
                        self.CODE = self.err_code[n]
                        self.OCC_TIME = self.create_file_name(order=1)
                        self.data_list = [self.linename, self.LOCATION, self.EVENT, self.CODE, self.OCC_TIME]
                        self.sand_data_list.append(self.data_list)
                        self.add_data_accdb(cnxn=self.cnxn, order=0)

                self.pre_event_type[n] = self.event_type[n]
                self.pre_err_code[n] = self.err_code[n]
                self.pre_tack_time[n] = self.tack_time[n]
    def read_resister(self):
        try:
            client = ModbusTcpClient(self.ip, self.port, timeout=3)
            data = client.read_holding_registers(self.device_error, len(self.location))
            data2 = client.read_holding_registers(self.device_operation, len(self.location))
            data3 = client.read_holding_registers(self.device_st, len(self.location))
            data4 = client.read_holding_registers(self.device_barcode, 20)
            data5 = client.read_holding_registers(self.device_plasma, 3)
            data = data.registers
            data2 = data2.registers
            data3 = data3.registers
            data4 = data4.registers
            data5 = data5.registers
            client.close()
        except Exception as err:
            print(err)
            print("Check ethernet cable")
            data = []
            data2 = []
            data3 = []
            data4 = [1]
            data5 = [1,1,1]
            for j in range(len(self.location)):
                data.append(0)
                data2.append(1)
                data3.append(1)
            return data, data2, data3, data4, data5
        return data, data2, data3, data4, data5

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = USER_window()

    sys.exit(app.exec_())