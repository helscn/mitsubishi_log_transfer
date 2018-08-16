# -*- coding: utf-8 -*-


import sys
import os
import csv
import struct
import socket
import pickle
from datetime import date, timedelta, datetime
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(750, 520)
        self.form = Form
        self.verticalLayout_7 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setSpacing(12)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setObjectName("groupBox")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.groupBox)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setContentsMargins(10, -1, -1, -1)
        self.horizontalLayout.setSpacing(23)
        self.horizontalLayout.setObjectName("horizontalLayout")

        self.areas = self.get_areas()
        self.chkAreas = []
        for area in self.areas:
            ck = QtWidgets.QCheckBox(self.groupBox)
            ck.setObjectName("chkArea" + area)
            ck.setToolTip("<html><head/><body><p>取消/选择 所有{0}区镭射机</p></body></html>".format(area))
            ck.setText(area + "区")
            self.horizontalLayout.addWidget(ck)
            self.chkAreas.append(ck)

        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.tblServers = QtWidgets.QTableWidget(self.groupBox)
        self.tblServers.setObjectName("tblServers")

        self.verticalLayout.addWidget(self.tblServers)
        self.verticalLayout_3.addLayout(self.verticalLayout)
        self.horizontalLayout_3.addWidget(self.groupBox)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.groupBox_2 = QtWidgets.QGroupBox(Form)
        self.groupBox_2.setObjectName("groupBox_2")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(self.groupBox_2)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setContentsMargins(5, 5, 5, -1)
        self.verticalLayout_4.setSpacing(10)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.label = QtWidgets.QLabel(self.groupBox_2)
        self.label.setObjectName("label")
        self.verticalLayout_4.addWidget(self.label)
        self.dateFrom = QtWidgets.QDateEdit(self.groupBox_2)
        self.dateFrom.setObjectName("dateFrom")
        self.verticalLayout_4.addWidget(self.dateFrom)
        self.label_2 = QtWidgets.QLabel(self.groupBox_2)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_4.addWidget(self.label_2)
        self.dateTo = QtWidgets.QDateEdit(self.groupBox_2)
        self.dateTo.setObjectName("dateTo")
        self.verticalLayout_4.addWidget(self.dateTo)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setContentsMargins(20, 0, 20, -1)
        self.horizontalLayout_2.setSpacing(4)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem1)
        self.label_3 = QtWidgets.QLabel(self.groupBox_2)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_2.addWidget(self.label_3)
        self.lcdTotalDays = QtWidgets.QLCDNumber(self.groupBox_2)
        self.lcdTotalDays.setObjectName("lcdTotalDays")
        self.horizontalLayout_2.addWidget(self.lcdTotalDays)
        self.label_4 = QtWidgets.QLabel(self.groupBox_2)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_2.addWidget(self.label_4)
        self.horizontalLayout_2.setStretch(2, 2)
        self.horizontalLayout_2.setStretch(3, 1)
        self.verticalLayout_4.addLayout(self.horizontalLayout_2)
        spacerItem2 = QtWidgets.QSpacerItem(20, 48, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_4.addItem(spacerItem2)
        self.verticalLayout_6.addLayout(self.verticalLayout_4)
        self.verticalLayout_2.addWidget(self.groupBox_2)
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setContentsMargins(10, -1, 10, 15)
        self.verticalLayout_5.setSpacing(20)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.btnCheckRemoteServers = QtWidgets.QPushButton(Form)
        self.btnCheckRemoteServers.setMinimumSize(QtCore.QSize(0, 23))
        self.btnCheckRemoteServers.setBaseSize(QtCore.QSize(0, 0))
        self.btnCheckRemoteServers.setObjectName("btnCheckRemoteServers")
        self.verticalLayout_5.addWidget(self.btnCheckRemoteServers)
        self.btnGetLog = QtWidgets.QPushButton(Form)
        self.btnGetLog.setObjectName("btnGetLog")
        self.verticalLayout_5.addWidget(self.btnGetLog)
        self.btnExit = QtWidgets.QPushButton(Form)
        self.btnExit.setObjectName("btnExit")
        self.verticalLayout_5.addWidget(self.btnExit)
        self.verticalLayout_2.addLayout(self.verticalLayout_5)
        self.horizontalLayout_3.addLayout(self.verticalLayout_2)
        self.horizontalLayout_3.setStretch(0, 3)
        self.horizontalLayout_3.setStretch(1, 1)
        self.verticalLayout_7.addLayout(self.horizontalLayout_3)
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setEnabled(True)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(False)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout_7.addWidget(self.progressBar)

        self.retranslateUi(Form)
        self.init_table(self.tblServers)
        self.init_date(self.dateFrom, self.dateTo)
        self.init_signal_connect()
        QtCore.QMetaObject.connectSlotsByName(Form)

    def checked_servers_count(self):
        count = 0
        for server in self.servers:
            if server['checkbox'].isChecked():
                count += 1
        return count

    def checkbox_changed(self):
        sender = self.form.sender()
        area = sender.text()[:1]
        for server in self.servers:
            if server['machine'].upper().find(area) >= 0:
                server['checkbox'].setCheckState(sender.checkState())

    def disable_button(self):
        self.btnCheckRemoteServers.setDisabled(True)
        self.btnGetLog.setDisabled(True)
        self.btnExit.setDisabled(True)

    def enable_button(self):
        self.btnCheckRemoteServers.setDisabled(False)
        self.btnGetLog.setDisabled(False)
        self.btnExit.setDisabled(False)

    def check_remote_servers(self):
        self.disable_button()
        failed_count = 0
        for i, server in enumerate(self.servers):
            try:
                sock = connect_to(server['ip'], server['port'])
                send_data(sock, False)
                ok_text = QtWidgets.QTableWidgetItem('连接正常')
                ok_text.setBackground(QtGui.QColor(200, 255, 200))
                self.tblServers.setItem(i, 4, ok_text)
                sock.close()
            except:
                failed_count += 1
                error_text = QtWidgets.QTableWidgetItem('无法连接')
                error_text.setBackground(QtGui.QColor(200, 50, 50))
                self.tblServers.setItem(i, 4, error_text)
            self.progressBar.setValue(int((i + 1) * 100 / len(self.servers)))
        self.enable_button()
        if failed_count == 0:
            QMessageBox.information(self.form, '检查完毕', '镭射机远程连接检查完成，所有镭射机连接正常。', QMessageBox.Ok)
        else:
            QMessageBox.warning(self.form, '检查完毕',
                                '镭射机远程连接检查完成，共有 {0} 台机无法连接！\n请检查镭射机上 LogServer 程序是否正常运行。'.format(str(failed_count)),
                                QMessageBox.Ok)
        self.progressBar.setValue(0)

    def get_servers(self):
        # 读取镭射机服务器信息文件并返回服务器列表
        filename = r'Servers.csv'
        servers = []
        try:
            with open(filename) as f:
                f_csv = csv.reader(f)
                for row in f_csv:
                    if len(row) == 3 and row[2].isdigit():
                        row[2] = int(row[2])
                        # servers.append(row)
                        servers.append({'machine': row[0], 'ip': row[1], 'port': row[2]})
        except:
            QMessageBox.warning(self.form, "错误", "无法读取服务器列表文件：\n" + filename, QMessageBox.Ok)
        return servers

    def get_areas(self):
        servers = self.get_servers()
        areas = []
        for server in servers:
            area = server['machine'][:1].upper()
            if area not in areas:
                areas.append(area)
        return areas

    def get_log_formats(self):
        filename = r'LogFormats.csv'
        log_formats = []
        try:
            with open(filename) as f:
                f_csv = csv.reader(f)
                for row in f_csv:
                    if len(row) >= 3 and (row[0].upper() in ['TRUE', '1', 'YES', 'V']):
                        log_formats.append({'remoteFilePath': row[1], 'localFileName': row[2]})
        except:
            QMessageBox.warning(self.form, "错误", "无法读取加工记录格式文件：\n" + filename, QMessageBox.Ok)
        return log_formats

    def ouput_Log(self):
        log_formats = self.get_log_formats()
        if len(log_formats) == 0:
            QMessageBox.warning(self.form, '警告', '加工记录格式文件中没有指定要读取的加工记录！', QMessageBox.Ok)
            return

        servers_count = self.checked_servers_count()
        if servers_count == 0:
            QMessageBox.information(self.form, '提示', '没有选择需要导出数据的镭射机！', QMessageBox.Ok)
            return

        save_path = QtWidgets.QFileDialog.getExistingDirectory(self.form, '选择加工记录保存路径', r'.')
        if not save_path:
            return

        self.disable_button()
        dtFrom = self.dateFrom.date().toPyDate()
        dtTo = self.dateTo.date().toPyDate()
        count = 0
        for i, server in enumerate(self.servers):
            self.progressBar.setValue(int(count * 100 / servers_count))
            if server['checkbox'].isChecked():
                failed = False
                try:
                    sock = connect_to(server['ip'], server['port'])
                except:
                    failed = True
                    error_text = QtWidgets.QTableWidgetItem('无法连接')
                    error_text.setBackground(QtGui.QColor(200, 50, 50))
                    self.tblServers.setItem(i, 4, error_text)
                    count += 1
                    continue
                for log in log_formats:
                    try:
                        with open(os.path.join(save_path, log['localFileName'].format(**server)), 'wt') as f:
                            get_header = False
                            for d in date_range(dtTo, dtFrom):
                                send_data(sock, log['remoteFilePath'].format(**d))
                                header, data = parse_log(recv_data(sock))
                                if not get_header and header:
                                    f.write(header + '\n')
                                    get_header = True
                                if data:
                                    f.write(data + '\n')
                    except:
                        failed = True
                        error_text = QtWidgets.QTableWidgetItem('无法保存记录')
                        error_text.setBackground(QtGui.QColor(200, 50, 50))
                        self.tblServers.setItem(i, 4, error_text)
                try:
                    send_data(sock, False)  # 通知远程服务器关闭连接
                    sock.close()
                except:
                    pass
                if not failed:
                    ok_text = QtWidgets.QTableWidgetItem('已导出')
                    ok_text.setBackground(QtGui.QColor(200, 255, 200))
                    self.tblServers.setItem(i, 4, ok_text)
                count += 1
        self.progressBar.setValue(int(count * 100 / servers_count))
        self.enable_button()
        QMessageBox.information(self.form, '导出完成', '所有选择的镭射机加工记录已导出完毕。', QMessageBox.Ok)
        self.progressBar.setValue(0)

    def update_days(self):
        self.dateFrom.setMaximumDate(self.dateTo.date())
        self.dateTo.setMinimumDate(self.dateFrom.date())
        self.lcdTotalDays.display(abs(self.dateTo.date().daysTo(self.dateFrom.date())) + 1)

    def init_table(self, table):
        # 初始化远程服务器表格清单
        table.setGeometry(QtCore.QRect(40, 100, 571, 291))
        table.setObjectName("tblServers")
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(['选择', '机台', 'IP', '端口', '状态'])
        table.horizontalHeader().setSectionsClickable(False)
        font = QtGui.QFont('宋体')
        font.setBold(True)
        table.horizontalHeader().setFont(font)  # 设置表头字体
        table.setColumnWidth(0, 50)
        table.setColumnWidth(1, 80)
        table.setColumnWidth(2, 120)
        table.setColumnWidth(3, 80)

        header = table.horizontalHeader()
        # header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        # header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        table.setRowCount(0)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        self.servers = self.get_servers()
        for server in self.servers:
            row = table.rowCount()
            table.setRowCount(row + 1)

            # 服务器列表中增加服务器表格中的Checkbox对象
            ck = QtWidgets.QCheckBox()
            server['checkbox'] = ck

            h = QtWidgets.QHBoxLayout()
            h.setAlignment(QtCore.Qt.AlignCenter)
            h.addWidget(ck)
            w = QtWidgets.QWidget()
            w.setLayout(h)

            table.setCellWidget(row, 0, w)
            table.setItem(row, 1, QtWidgets.QTableWidgetItem(server['machine']))
            table.setItem(row, 2, QtWidgets.QTableWidgetItem(server['ip']))
            table.setItem(row, 3, QtWidgets.QTableWidgetItem(str(server['port'])))

    def init_date(self, *dates):
        for d in dates:
            d.setDate(QtCore.QDate.currentDate())
            d.setMaximumDate(QtCore.QDate.currentDate())
        self.lcdTotalDays.display(1)

    def init_signal_connect(self):
        for ck in self.chkAreas:
            ck.stateChanged.connect(self.checkbox_changed)
        self.dateFrom.dateChanged.connect(self.update_days)
        self.dateTo.dateChanged.connect(self.update_days)
        self.btnCheckRemoteServers.clicked.connect(self.check_remote_servers)
        self.btnGetLog.clicked.connect(self.ouput_Log)
        self.btnExit.clicked.connect(QtCore.QCoreApplication.instance().quit)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "三菱镭射机加工记录导出程序"))
        self.groupBox.setTitle(_translate("Form", "镭射机台选择"))
        self.groupBox_2.setTitle(_translate("Form", "加工时间"))
        self.label.setText(_translate("Form", "起始时间："))
        self.label_2.setText(_translate("Form", "结束时间："))
        self.label_3.setText(_translate("Form", "共"))
        self.label_4.setText(_translate("Form", "天"))
        self.btnCheckRemoteServers.setToolTip(
            _translate("Form", "<html><head/><body><p>检查所有镭射机的远程连接状态，无法连接的镭射机将无法导出加工记录。</p></body></html>"))
        self.btnCheckRemoteServers.setText(_translate("Form", "检查连接状态"))
        self.btnGetLog.setToolTip(_translate("Form", "<html><head/><body><p>导出所有选择的镭射机加工记录。</p></body></html>"))
        self.btnGetLog.setText(_translate("Form", "导出加工记录"))
        self.btnExit.setToolTip(_translate("Form", "<html><head/><body><p>关闭镭射机加工记录导出程序。</p></body></html>"))
        self.btnExit.setText(_translate("Form", "关闭程序"))


def connect_to(host, port):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(3)  # 指定Socket连接超时时间
    sock.connect((host, port))
    return sock


def send_data(the_socket, data):
    data = pickle.dumps(data, True)
    the_socket.sendall(struct.pack('>I', len(data)) + data)


def recv_data(the_socket):
    sock_data = bytes()  # 通过Socket缓冲区获取的临时数据
    recv_size = 0  # 已经接收的数据总大小
    total_data = []  # 保存所有接收数据的数组
    total_size = sys.maxsize  # 需要接收的总数据大小，保存在数据头的前4个字节
    get_data_size = False  # 是否已经解析数据头的大小数据

    while recv_size < total_size:
        sock_data = the_socket.recv(8192)
        recv_size += len(sock_data)
        if get_data_size:
            total_data.append(sock_data)
        else:
            if len(total_data) == 0:
                total_data.append(sock_data)
            else:
                total_data[0] += sock_data
            if len(total_data[0]) >= 4:
                total_size = struct.unpack('>I', total_data[0][:4])[0]
                total_data[0] = total_data[0][4:]
                recv_size -= 4
                get_data_size = True
    return pickle.loads(b''.join(total_data))


def parse_log(log_data, coding='ANSI'):
    if not log_data:
        return ('', '')
    if type(log_data) is bytes:
        log_data = log_data.decode(coding)
    log = log_data.splitlines()
    log = [line for line in log if len(line) > 0]
    header = ''
    if len(log) > 0:
        header = log[0]
        log = log[1:]
    log.reverse()
    return (header, '\n'.join(log))


def date_range(date1, date2):
    # 迭代返回指定范围内的日期数据
    step = timedelta(days=1) if date1 <= date2 else timedelta(days=-1)
    while date1 - date2 != step:
        yield {'Y': date1.strftime('%Y'),
               'y': date1.strftime('%y'),
               'm': date1.strftime('%m'),
               'd': date1.strftime('%d'),
               'date': date1
               }

        date1 += step


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
