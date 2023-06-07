# -- coding: utf-8 --
# @info:
# @Author : liyahui
# @Time : 2023/5/31 上午10:33
# @File : UPS3072Adapter.py
# @Software: PyCharm
"""通过串口发送测试指令，通过modbus读取数据并解析，主要参数为电池电压，电流，电阻，温度，功率"""
import datetime
import queue
import threading

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, GradientFill, Border, colors, Side
from tqdm import tqdm
import random
import time

import serial
import modbus_tk.defines as cst
from modbus_tk import modbus_rtu
import serial.tools.list_ports
import configparser

from LogFormat import LogFormat

log = LogFormat('UPS3072Adapter', "UPS3072Adapter.log")
logger = log.logger



class UPS3072Adapter(object):
    def __init__(self, file_path):
        self.temp_range = None
        self.v_range = None
        self.config_path = file_path
        self.port = None
        self.addr = None
        # 波特率
        self.baud_rate = None
        self.timeout = None
        self.wait_time_out = None
        # 字节大小
        self.bytesize = None
        # 校验位N－无校验，E－偶校验，O－奇校验
        self.parity = None
        # 停止位
        self.stop_bits = None
        self.modbus_rtu_obj = None
        self.addr_mapping = None
        self.start_addr = None
        self.table_head = None
        self.nodes_list = None
        self.log_q = queue.Queue()
        self.init_conf()

    @staticmethod
    def _version():
        """return: 版本号"""
        return '1.1.0'

    def init_conf(self):
        """加载配置文件，初始化配置信息"""
        config = configparser.ConfigParser()
        try:
            config.read(self.config_path, encoding='utf-8')
            self.port = config.get('UPS3072Adapter', 'port')
            self.baud_rate = config.getint('UPS3072Adapter', 'baud_rate')
            self.addr = config.getint('UPS3072Adapter', 'addr')
            self.stop_bits = config.getint('UPS3072Adapter', 'stop_bits')
            self.parity = config.get('UPS3072Adapter', 'parity')
            self.start_addr = config.getint('UPS3072Adapter', 'start_addr')
            self.bytesize = config.getint('UPS3072Adapter', 'bytesize')
            self.timeout = config.getint('UPS3072Adapter', 'timeout')
            self.addr_mapping = eval(config.get('UPS3072Adapter', 'addr_mapping'))
            self.table_head = eval(config.get('UPS3072Adapter', 'table_head'))
            self.nodes_list = eval(config.get('UPS3072Adapter', 'nodes_list'))
            self.v_range = config.getint('UPS3072Adapter', 'v_range')
            self.temp_range = config.getint('UPS3072Adapter', 'temp_range')
            self.wait_time_out = config.getint('UPS3072Adapter', 'wait_time_out')
        except Exception as e:
            error_msg = f'加载配置发生错误：{e}'
            logger.error(error_msg)
            self.log_q.put(error_msg)
            return False, error_msg

    def connect(self):
        """建立连接"""
        try:
            # 设定串口为从站
            ser = serial.Serial(self.port, self.baud_rate, self.bytesize, parity=self.parity, stopbits=self.stop_bits)
            self.modbus_rtu_obj = modbus_rtu.RtuMaster(ser)
            self.modbus_rtu_obj.set_timeout(self.timeout)
            log_msg = '3702UPS system connect success'
            self.log_q.put(log_msg)
            logger.info(log_msg)
            return True, log_msg
        except Exception as e:
            log_msg = '3702UPS system connect failed message %s' % str(e)
            logger.error(log_msg)
            self.log_q.put(log_msg)
            return False, log_msg

    def read_hold_register(self, excel_file_path, time_str=None, sn=None):
        """
        读保持寄存器 : 03H
        :param excel_file_path:
        :param time_str: 测试时间字符串
        :param sn: 被测设备唯一标识符
        :return: (Bool,data) 成功/失败，成功/失败提示信息
        """
        th = threading.Thread(target=self.write_number)
        th.start()
        time.sleep(2)
        if not time_str:
            time_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        length = len(list(self.addr_mapping.keys()))
        start_addr = self.start_addr
        real_value = self.table_head[3]
        self.table_head.remove(real_value)
        self.nodes_list.sort()
        try:
            new_list = []
            if len(self.nodes_list) == 0:
                log_msg = f"请配置检测节点"
                logger.info(log_msg)
                self.log_q.put(log_msg)
                return False,log_msg
            for node_index, node_value in tqdm(enumerate(self.nodes_list)):
                data = []
                start_time = datetime.datetime.now()
                end_time = start_time + datetime.timedelta(minutes=self.wait_time_out)
                current_time = datetime.datetime.now()
                log_msg = f"当前期望节点值:{float(node_value)}"
                logger.info(log_msg)
                self.log_q.put(log_msg)
                return_data = self.modbus_rtu_obj.execute(self.addr, cst.READ_HOLDING_REGISTERS, start_addr, length)
                real_value = round(float(return_data[35] / 1000), 1)

                while current_time <= end_time:
                    return_data = self.modbus_rtu_obj.execute(self.addr, cst.READ_HOLDING_REGISTERS, start_addr, length)
                    real_value = round(float(return_data[35] / 1000), 2)
                    if real_value > float(node_value):
                        data = return_data
                        log_msg = f"当前电压值:{real_value},期望的节点值:{node_value}，实际值高于期望值，将直接读取此时数据,然后等待下一个节点"
                        logger.info(log_msg)
                        self.log_q.put(log_msg)
                        break
                    elif real_value == float(node_value):
                        data = return_data
                        log_msg = f"当前电压值:{real_value},期望的节点值:{node_value},实际值等于期望值，将开始读取此时数据"
                        logger.info(log_msg)
                        self.log_q.put(log_msg)
                        break
                    else:
                        current_time = datetime.datetime.now()
                        log_msg = f'当前电压值:{real_value},期望的节点值:{node_value},等待截至时间:{end_time},waiting......'
                        logger.info(log_msg)
                        self.log_q.put(log_msg)
                        time.sleep(10)
                else:
                    log_msg = f'等待超时，测试终止！'
                    logger.info(log_msg)
                    self.log_q.put(log_msg)
                    return False, log_msg
                log_msg = f"开始读取第{node_index + 1}节点"
                logger.info(log_msg)
                self.log_q.put(log_msg)
                v_sum = float(sum(data[35:52]))  # 电压之和
                v_menu = round(float(sum(data[35:52]) / 16000), 3)  # 电压平均值
                temp_menu = round(float(sum(data[52:]) / 160), 1)  # 温度平均值
                standar_str = f"{real_value}({node_value})\n[标准值:电压：{v_menu}±{self.v_range}V;温度：{temp_menu}±{self.temp_range}℃]]"
                self.table_head.insert(3 + node_index, standar_str)
                if node_index == 0:  # 第一个节点，从配置文件里读取数据列表，并填充值
                    for data_index, item in enumerate(data):
                        code = data_index + 1
                        dict_key = str(data_index + self.start_addr)
                        value_list = self.addr_mapping.get(dict_key, [])
                        add_items = len(self.nodes_list)
                        for i in range(0, add_items - 1):  # 先把所有列补出来
                            value_list.insert(-2, "")  # 序号","测试项目","方法","实测值1","实测值2","实测值3","结果判断","描述
                        value_list.insert(0, code)
                        new_list.append(value_list)

                for data_index, item in enumerate(data):
                    code = data_index + 1
                    if code == 35:  # 总电压
                        item = round(float(item / 1000), 3)
                        new_list[data_index][3 + node_index] = item
                        if abs(item - v_sum) <= self.v_range:
                            new_list[data_index][-2] = "合格"
                        else:
                            new_list[data_index][-2] = "不合格"
                    elif 35 <= code <= 51:  # 电压
                        item = round(float(item / 1000), 3)
                        new_list[data_index][3 + node_index] = item
                        if abs(item - v_menu) <= self.v_range:
                            new_list[data_index][-2] = "合格"
                        else:
                            new_list[data_index][-2] = "不合格"
                    elif 52 <= code <= 67:  # 温度
                        item = round(float(item / 10), 1)
                        new_list[data_index][3 + node_index] = item
                        if abs(item - temp_menu) <= self.temp_range:
                            new_list[data_index][-2] = "合格"
                        else:
                            new_list[data_index][-2] = "不合格"
                    else:  # 其他
                        new_list[data_index][3 + node_index] = item
                log_msg = f"第{node_index + 1}个节点数据已读取,{new_list}"
                logger.info(log_msg)
                self.log_q.put(log_msg)
            log_msg = f"测试完成,获取到测试结果:{new_list},"
            logger.info(log_msg)
            self.log_q.put(log_msg)
        except Exception as e:
            log_msg = '测试出错: %s' % str(e)
            logger.error(log_msg)
            self.log_q.put(log_msg)
            return False, log_msg
        else:
            try:
                self.write_to_file(self.table_head, new_list, excel_file_path, time_str, sn)
                log_msg += "生成报告成功"
                logger.info(log_msg)
                self.log_q.put(log_msg)
                return True, log_msg
            except Exception as e:
                log_msg += f"生成报告出错:{e}"
                logger.info(log_msg)
                self.log_q.put(log_msg)
                return False, log_msg

    def write_number(self):
        for node_index, node_value in enumerate(self.nodes_list):
            logger.info(f"写入第{node_index + 1}节点")
            for j in range(0, 70):
                random_num = 1
                if node_index == 0:
                    random_num = random.randint(3200, 3220)
                elif node_index == 1:
                    random_num = random.randint(3300, 3320)
                elif node_index == 2:
                    random_num = random.randint(3400, 3420)
                elif node_index == 3:
                    random_num = random.randint(3500, 3520)
                self.modbus_rtu_obj.execute(1, cst.WRITE_SINGLE_REGISTER, j, output_value=random_num)
            time.sleep(10)

    def write_to_file(self,table_head, data_list, excel_file_path, time_str, sn=None):
        """写入excel文件"""
        wb = openpyxl.Workbook()
        sheet = wb.active
        table_head_font = Font(
            name="微软雅黑",  # 字体
            size=16,  # 字体大小
            bold=True,  # 是否加粗，True/False
            italic=False,  # 是否斜体，True/False
            strike=None,  # 是否使用删除线，True/False
            underline=None,  # 下划线, 可选'singleAccounting', 'double', 'single', 'doubleAccounting'
        )
        # 自定义字体样式
        font = Font(
            name="微软雅黑",  # 字体
            size=13,  # 字体大小
            bold=False,  # 是否加粗，True/False
            italic=False,  # 是否斜体，True/False
            strike=None,  # 是否使用删除线，True/False
            underline=None,  # 下划线, 可选'singleAccounting', 'double', 'single', 'doubleAccounting'
        )

        fill = PatternFill(
            patternType="solid",  # 填充类型，可选none、solid、darkGray、mediumGray、lightGray、lightDown、lightGray、lightGrid
            fgColor="99CCFF",  # 前景色，16进制rgb
            bgColor="99CCFF",  # 背景色，16进制rgb
            # fill_type=None,     # 填充类型
            # start_color=None,   # 前景色，16进制rgb
            # end_color=None      # 背景色，16进制rgb
        )
        border = Border(top=Side(border_style="thin", color=colors.BLACK),
                        bottom=Side(border_style="thin", color=colors.BLACK),
                        left=Side(border_style="thin", color=colors.BLACK),
                        right=Side(border_style="thin", color=colors.BLACK))
        table_border = Border(top=Side(border_style="medium", color=colors.BLACK),
                              bottom=Side(border_style="medium", color=colors.BLACK),
                              left=Side(border_style="medium", color=colors.BLACK),
                              right=Side(border_style="medium", color=colors.BLACK))
        sheet.column_dimensions["A"].width = 10
        sheet.column_dimensions["B"].width = 30
        sheet.column_dimensions["C"].width = 20
        sheet.column_dimensions["D"].width = 30
        sheet.column_dimensions["E"].width = 30
        sheet.column_dimensions["F"].width = 30
        sheet.column_dimensions["G"].width = 20
        sheet.column_dimensions["H"].width = 100
        sheet.cell(1, 1).value = "SN:"
        sheet.cell(1, 1).font = table_head_font
        sheet.cell(1, 2).value = sn
        sheet.cell(1, 2).font = table_head_font
        sheet.cell(1, 3).value = "测试日期:"
        sheet.cell(1, 3).font = table_head_font
        sheet.cell(1, 4).value = time_str
        sheet.cell(1, 4).font = table_head_font

        for index, item in enumerate(table_head):
            cell = sheet.cell(row=2, column=index + 1)
            sheet.cell(row=1, column=index + 1).fill = fill
            sheet.cell(row=1, column=index + 1).border = table_border
            cell.value = item
            cell.font = table_head_font
            cell.fill = fill
            cell.border = table_border
            cell.alignment = Alignment(wrapText=True)

        for row_index, data_item in enumerate(data_list):
            for col_index, data in enumerate(data_item):
                cell = sheet.cell(row=row_index + 3, column=col_index + 1)
                cell.value = data
                cell.font = font
                cell.border = border
        wb.save(excel_file_path)
        log_msg = f"写入Excel文件：{excel_file_path}"
        logger.info(log_msg)
        self.log_q.put(log_msg)


if __name__ == "__main__":
    a = UPS3072Adapter("./conf/UPS3072Adapter.conf")
    a.connect()
    a.read_hold_register("./data.xlsx")
