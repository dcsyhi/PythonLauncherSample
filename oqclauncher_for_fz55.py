#!/usr/bin/env python
# -*- coding: Shift-JIS -*-

from abc import ABCMeta
from abc import ABCMeta, abstractmethod
import dataclasses
import configparser
import datetime
import errno
import os
import pathlib
import subprocess
import ctypes, sys
import sys
import tkinter as tk
from tkinter import Menu, messagebox, ttk
import tkinter.filedialog as tkf
import tkinter.messagebox as tkm
from tkinter.scrolledtext import ScrolledText
from typing import NoReturn, Type
import time

import openpyxl
import pandas as pd

# �f�[�^���i�[����N���X�̐錾
@dataclasses.dataclass
class Module:
    name: str
    val: int = 0

@dataclasses.dataclass
class Passed:
    name: str
    val: int = 1
   
@dataclasses.dataclass
class Sum:
    name: str
    val: int = 0

@dataclasses.dataclass
class Model:
    sku: str = ""
    serial: str = ""



# �e��OQC���ڂ̃t���O��錾����
wwan_flg     = Module('WWAN_FLG')    # WWAN
wlan2nd_flg  = Module('WLAN2ND_FLG') # 2nd WLAN
gps_flg      = Module('GPS_FLG')     # GPS
dp_flg       = Module('DP_FLG')      # Dual Pass
cam_flg      = Module('CAM_FLG')     # IR Camera
finger_flg   = Module('FINGER_FLG')  # Finger Print
rfid_flg     = Module('RFID_FLG')    # NFC RFID
scr_flg      = Module('SCR_FLG')     # Smart Card Reader
batt2nd_flg  = Module('BATT2ND_FLG') # 2nd Battery
vga_flg      = Module('VGA_FLG')     # VGA
seri_flg     = Module('SERI_FLG')    # Serial
usb3_flg     = Module('USB3_FLG')    # USB 3.0
lan2nd_flg   = Module('LAN2ND_FLG')  # 2nd LAN
rgg_usb_flg  = Module('RGG_USB_FLG') # Rugged USB
odd_flg      = Module('ODD_FLG')     # ODD
dgpu_flg     = Module('DGPU_FLG')    # dGPU
scr2_flg     = Module('SCR2_FLG')    # Smart Card Reader 2
ssd2nd_flg   = Module('SSD2ND_FLG')  # 2nd SSD
nonbkl_flg   = Module('NONBKL_FLG')  # non Backlit KBD
bkl_flg      = Module('BKL_FLG')     # Backlit KBD
wlan_flg     = Module('WLAN_FLG')    # WLAN

# OK,NG����t���O
passed_flg   = Passed('Passed')

# OK���肪�o�����ڂ̍��v�l
sum = Sum('Sum')

# �@�햼�A�V���A������ێ�����
model = Model()

# �\�t�g�E�F�A�̎��s(�J�n)
root = tk.Tk()

# --------------------------------------------------
# config.ini�ɂ��o�b�`���{���ڂ̐ݒ�
# --------------------------------------------------
config_ini = configparser.ConfigParser()
config_ini_path = './batch/config_fz55.ini'

# ini�t�@�C�������݂��邩�`�F�b�N
if os.path.exists(config_ini_path):
    # ini�t�@�C�������݂���ꍇ�A�t�@�C����ǂݍ���
    with open(config_ini_path, encoding='utf-8') as fp:
        config_ini.read_file(fp)
    # ini�t�@�C���ɋL�ڂ����A�o�b�`���s�p�X�̒l���擾
    read_default   = config_ini['FZ-55']
    cam_folder     = read_default.get('IRCamera_Folder')
    finger_folder  = read_default.get('FingerPrint_Folder')
    rfid_folder    = read_default.get('RFID_Folder')
    scr_folder     = read_default.get('SCR_Folder')
    scr2_folder    = read_default.get('SCR2_Folder')
    dgpu_folder    = read_default.get('dGPU_Folder')

    cam_name     = read_default.get('IRCamera_File')
    finger_name  = read_default.get('FingerPrint_File')
    rfid_name    = read_default.get('RFID_File')
    scr_name     = read_default.get('SCR_File')
    scr2_name    = read_default.get('SCR2_File')
    dgpu_name    = read_default.get('dGPU_File')

else:
    # ini�t�@�C�������݂��Ȃ��ꍇ�A�G���[����
    raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), config_ini_path)

# �O�̂��߁A���O�Ƀt�@�C�����폜���Ă���
subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\sku.txt', shell=True)
subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\serial.txt', shell=True)
subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\eco_num.txt', shell=True)
subprocess.run(r'fld_accs\existdel.wsf -DEL=.\output\aim_data.csv', shell=True)


# ------------------------------
# ��ʂ̏���������ѕϐ��̏�����
# ------------------------------
def initialize():
    """
    ��ʂ̏���������ѕϐ��̏����������s����
    -----------------
    �ȉ��̒l�͂��ׂ�"0"�ɐݒ肷��
    wwan_flg #WWAN: int
    wlan2nd_flg #2nd WLAN : int
    gps_flg  #GPS : int
    dp_flg   #Dual Pass : int
    cam_flg  #IR Camera : int
    finger_flg  #Finger Print : int
    rfid_flg  #NFC RFID : int
    scr_flg   #Smart Card Reader : int
    batt2nd_flg  #2nd Battery : int
    vga_flg  #VGA : int
    seri_flg #Serial : int
    usb3_flg #USB 3.0 : int
    lan2nd_flg  #2nd LAN : int
    rgg_usb_flg  #Rugged USB : int
    odd_flg   #ODD : int
    dgpu_flg  #dGPU : int
    scr2_flg  #Smart Card Reader 2 : int
    ssd2nd_flg #2nd SSD : int
    nonbkl_flg #non Backlit KBD : int
    bkl_flg  #Backlit KBD : int
    wlan_flg #WLAN : int
       
    passed_flg : int, ���{�K�v��OQC���ڂ����ׂĎ��s���ꂽ���𔻒肷��B�����l��"1"�ɐݒ�

    """
    
    # �������ݒ�
    # TODO:�t���O�͋@��ˑ�
    wwan_flg.val     = 0    # WWAN
    wlan2nd_flg.val  = 0    # 2nd WLAN
    gps_flg.val      = 0    # GPS
    dp_flg.val       = 0    # Dual Pass
    cam_flg.val      = 0    # IR Camera
    finger_flg.val   = 0    # Finger Print
    rfid_flg.val     = 0    # NFC RFID
    scr_flg.val      = 0    # Smart Card Reader
    batt2nd_flg.val  = 0    # 2nd Battery
    vga_flg.val      = 0    # VGA
    seri_flg.val     = 0    # Serial
    usb3_flg.val     = 0    # USB 3.0
    lan2nd_flg.val   = 0    # 2nd LAN
    rgg_usb_flg.val  = 0    # Rugged USB
    odd_flg.val      = 0    # ODD
    dgpu_flg.val     = 0    # dGPU
    scr2_flg.val     = 0    # Smart Card Reader 2
    ssd2nd_flg.val   = 0    # 2nd SSD
    nonbkl_flg.val   = 0    # non Backlit KBD
    bkl_flg.val      = 0    # Backlit KBD
    wlan_flg.val     = 0    # WLAN

    # OQC��"Pass"�����ǂ����𔻒肷��ϐ���1�ɐݒ肵�Ă���
    passed_flg.val   = 1

    # �����t�@�C�����c���Ă�����폜����
    subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\sku.txt', shell=True)
    subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\serial.txt', shell=True)
    subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\eco_num.txt', shell=True)
    subprocess.run(r'fld_accs\existdel.wsf -DEL=.\output\aim_data.csv', shell=True)


# --------------------------------------------------------
# �`�F�b�N�V�[�g�ǂݍ��݁`�E�B�W�F�b�g�`��i���C���֐��ɑ����j
# ---------------------------------------------------------
def dispOQCcontents():
    """
    �`�F�b�N�V�[�g�ǂݍ��݁`�E�B�W�F�b�g�`����s���i���C���֐��ɑ����j

    *    �`�F�b�N�V�[�g�̃t�@�C���p�X���w��
    * -> �`�F�b�N����@���BCR�œǂݍ���(�c�[����CharToF6.exe���g�p) 
    * -> �����`����ʂ̃{�^���F��������
    * -> SKU, Revision����ʂɕ\������
    * -> ���O�Ɏ��{����ǋL
    * -> �`�F�b�N�V�[�g����2�����`����DataFrame�Ƃ��Ē��o
    * -> ���o�f�[�^��.csv�t�@�C��(aim_data.csv)�ɏo�� 
    * -> ���{�K�v�ȃ{�^���F��"��"�F�ɕύX(���{�s�v��"�D"�F�ɕύX����)
    * -> �����`�F�b�N�V�[�g�̂��ׂĂ̍��ڂ�"��"�������Ă��Ȃ������烍�O���o�͂����Ȃ�

    Parameters
    ----------
    folder_name : str
        �`�F�b�N�V�[�g��u���t�H���_�̃p�X���i�[
    file_name : str
        ��Lfolder_name�ɏ����ꂽ�t�H���_�̒�����A"JIGSSD"��ړ����Ƃ��Ď��t�@�C���p�X���i�[    
    txt : str
        SKU���A�V���A�������i�[
    wb : Workbook
        ���[�N�u�b�N�̎w��(openpyxl)
    net_err : int
        net use�R�}���h�Ńl�b�g���[�N�h���C�u�Ƃ��Đڑ�����ۂ̕Ԃ�l���i�[����
    
    Inner Functions
    ---------------
    read_excel(����1)
        - 2�����z��(DataFrame)�̎擾 ����у{�^���\���F�̕ύX�B����1��'Matrix(EN)'���w��
        read_excel���̃l�X�g���ꂽ�֐��͈ȉ��̒ʂ�
            - get_value_list
            - get_dataframe
            - set_bg_color 

    """

    # �e��OQC���ڂ̃t���O������������"initialibze"�֐������s
    initialize()
    
    # net use�R�}���h�Ńl�b�g���[�N�h���C�u�Ƃ��Đڑ�����
    subprocess.run(r'Net Use X: /delete /y', shell=True)
    try:
        # net_err = subprocess.run(r'Net Use X: \\172.24.3.15\Imaging\QALOG\HCS_Config H3@rtlandCSConfig /user:HCS_TestApp /Persistent:No', shell=True, timeout=5)
        net_err = subprocess.run(r'Net Use X: \\132.182.76.44\Imaging\QALOG\PCI_Config /Persistent:No', shell=True, timeout=5)
    except subprocess.TimeoutExpired:
        messagebox.showerror('Error', 'Cannot found Network. Please ask manager.')
        sys.exit(1)
    

    # �`�F�b�N�V�[�g�̃t�@�C���p�X���w��
    folder_name: str = r"X:\Templates"

    # folder_name�ɏ����ꂽ�t�H���_�̒�����AJIGSSD*.xlsx����������file_name�Ƀt�@�C���p�X�Ƃ��ĕ�������i�[����
    file_names: list = list(pathlib.Path(folder_name).glob('JIGSSD*.xlsx'))
    print(file_names)
    if file_names == []:
        messagebox.showerror('Error', 'Cannot found matrix sheet. Please ask manager.')
        raise IndexError
    file_name = str(file_names[0])
    print(file_name)

    # �`�F�b�N����@���BCR�œǂݍ���(�c�[����CharToF6.exe���g�p) 
    subprocess.run(r'kb\CharToF6.exe /M"Config Model Name" /f.\exe\sku.txt', shell=True)
    chk_eco = subprocess.run(r'kb\MessBtn6.exe /M"Is this ECO config ?', shell=True).returncode
    if chk_eco == 0:
        subprocess.run(r'kb\CharToF6.exe /M"ECO Number ?" /f.\exe\eco_num.txt', shell=True)
    elif chk_eco == 1:
        pass
    subprocess.run(r'kb\CharToF6.exe /M"Config Serial Number" /f.\exe\serial.txt', shell=True)

    # �����`����ʂ̃{�^���F��������
    Button1.configure(fg='black',bg='SystemButtonFace')  #WWAN
    Button2.configure(fg='black',bg='gray')              #2nd WLAN
    Button3.configure(fg='black',bg='SystemButtonFace')  #GPS
    Button4.configure(fg='black',bg='gray')              #Dual Pass
    Button5.configure(fg='black',bg='SystemButtonFace')  #IR Camera
    Button6.configure(fg='black',bg='SystemButtonFace')  #Finger Print
    Button7.configure(fg='black',bg='SystemButtonFace')  #NFC RFID
    Button8.configure(fg='black',bg='SystemButtonFace')  #Smart Card Reader
    Button9.configure(fg='black',bg='gray')              #2nd Battery
    Button10.configure(fg='black',bg='SystemButtonFace') #VGA
    Button11.configure(fg='black',bg='SystemButtonFace') #Serial
    Button12.configure(fg='black',bg='SystemButtonFace') #USB 3.0
    Button13.configure(fg='black',bg='SystemButtonFace') #2nd LAN
    Button14.configure(fg='black',bg='SystemButtonFace') #Rugged USB
    Button15.configure(fg='black',bg='SystemButtonFace') #ODD
    Button16.configure(fg='black',bg='SystemButtonFace') #dGPU
    Button17.configure(fg='black',bg='SystemButtonFace') #Smart Card Reader 2
    Button18.configure(fg='black',bg='SystemButtonFace') #2nd SSD
    Button19.configure(fg='black',bg='SystemButtonFace') #non Backlit KBD
    Button20.configure(fg='black',bg='SystemButtonFace') #Backlit KBD
    Button21.configure(fg='black',bg='SystemButtonFace') #WLAN
    

    # �@��i�Ԃ���ʂɕ\������
    with open(r'.\exe\sku.txt', encoding='utf-8') as fp:
        model.sku = fp.readline().rstrip('\n')
        # ���x���̐���(BCR�œǂݎ����SKU��\��)
    try:
        with open(r'.\exe\eco_num.txt', encoding='utf-8') as fp:
            eco_num = '_' + fp.readline().rstrip('\n')
    except:
        eco_num = ''
    model.sku += eco_num
    label = tk.Label(text='SKU : {}'.format(model.sku))
    label.place(x=50, y=15)

    # �V���A������ʂɕ\������
    with open(r'.\exe\serial.txt', encoding='utf-8') as fp:
        model.serial = fp.readline().rstrip('\n')
        # ���x���̐���(BCR�œǂݎ����SKU��\��)
        label = tk.Label(text='SERIAL : {}'.format(model.serial))
        label.place(x=280, y=15)

    # �A�v���̃o�[�W��������ʂɕ\������
    with open(r'.\ver\fz55_version.txt', encoding='utf-8') as fp:
        version = fp.readline().rstrip('\n')
        # ���x���̐���(BCR�œǂݎ����SKU��\��)
        label = tk.Label(text='Application Version : {}'.format(version))
        label.place(x=510, y=510)
     

    # ���O�Ɏ��{����ǋL
    dt_now = datetime.datetime.now().strftime("%Y-%m-%d")
    txt.configure(state='normal')
    txt.insert('end', '[SKU:{0}/SERIAL:{1}]\n'.format(model.sku, model.serial))
    txt.insert('end', dt_now + '\n')
    txt.configure(state='disabled')


    # �`�F�b�N�V�[�g����2�����`����DataFrame�Ƃ��Ē��o
    # ���o�f�[�^��.csv�t�@�C��(aim_data.csv)�ɏo�� 
    wb = openpyxl.load_workbook(file_name, data_only=True)
    
    
    # 2�����z��(DataFrame)�̎擾(�����֐�)
    def read_excel(sheet_name: str):
        sheet = wb[sheet_name]
        # print(sheet_name)

        # �Z������Revision��ǂݍ���ŉ�ʂɕ\��������
        data_revision = sheet["P2"].value
        label2 = tk.Label(text='Sheet Revision : {}'.format(data_revision))
        label2.place(x=500, y=15)
        
       # 2�����z��(DataFrame)�̎擾(�����֐�)
        def get_value_list(t_2d: list):
            return([[cell.value for cell in row] for row in t_2d])

        def get_dataframe():
            l_2d: list = get_value_list(sheet['C29':'BN9000'])
            return l_2d

        df = pd.DataFrame(get_dataframe())
        # ���o�����f�[�^��.csv�t�@�C���ɏo�͂���
        df= df.dropna(how='all').dropna(how='all', axis=1)
        # TODO: Column���͋@��ˑ�
        df.columns = ["Config Model Name","Base Model Name","BIOS Model Name","DPK","SAR setting","WWAN ID","WLAN ID","WWAN SKU","IMEI No.",
                    "","","Kit number","S/N","","","Model name","Serial number","BIOS","EC","1st SSD Pack","1st SSD Pack","1st SSD Pack","1st SSD Pack",
                    "1st SSD Pack","1st SSD Pack","1st SSD Pack","2nd DIMM check","2nd DIMM check",
                    "WWAN","2nd WLAN","GPS","Dual Pass with dGPS","Dual Pass with dGPS","Dual Pass with dGPS","Dual Pass with dGPS","Dual Pass w/o dGPS","Dual Pass w/o dGPS",
                    "IR Camera","Finger Print","NFC RFID","Smart Card Reader","2nd Battery","VGA","VGA","VGA","Serial","Serial","Serial","USB3.0","2nd LAN","Rugged USB",
                    "DVD Multi","Blu-ray","dGPU","Smart Card Reader2","2nd SSD","2nd SSD","2nd SSD","2nd SSD","2nd SSD","non Backlit KBD","Backlit KBD","WLAN"]
        df = df.set_index('Config Model Name')
        data = df.iloc[0:0]
        data  = df.loc[model.sku]
        data.to_csv("output/aim_data.csv", header=True, index=True, encoding='utf_8_sig')
        data = pd.read_csv("output/aim_data.csv", index_col=0)  
        # print(data)
        data2 = data[data[model.sku] == "��"]
        # print(data2)
        d_index = data2.index.tolist()
        # print(d_index)

        # ���{�K�v�ȃ{�^���F��"��"�F�ɕύX 
        def set_bg_color(color: str):

            if "WWAN" in d_index :
                Button1.configure(fg='white', bg=color)
                wwan_flg.val = 1 # WWAN
            if "GPS" in d_index :
                Button3.configure(fg='white', bg=color)
                gps_flg.val = 1 # GPS
            if "IR Camera" in d_index :
                Button5.configure(fg='white', bg=color)
                cam_flg.val = 1 # IR Camera
            if "Finger Print" in d_index :
                Button6.configure(fg='white', bg=color)
                finger_flg.val = 1 # Finger Print
            if "NFC RFID" in d_index :
                Button7.configure(fg='white', bg=color)
                rfid_flg.val = 1 # NFC RFID
            if "Smart Card Reader" in d_index :
                Button8.configure(fg='white', bg=color)
                scr_flg.val = 1  # Smart Card Reader
            if "VGA" in d_index :
                Button10.configure(fg='white', bg=color)
                vga_flg.val = 1 # VGA
            if "Serial" in d_index :
                Button11.configure(fg='white', bg=color)
                seri_flg.val = 1 # Serial
            if "USB3.0" in d_index :
                Button12.configure(fg='white', bg=color)
                usb3_flg.val = 1  # USB 3.0
            if "2nd LAN" in d_index :
                Button13.configure(fg='white', bg=color)
                lan2nd_flg.val  = 1 # 2nd LAN
            if "Rugged USB" in d_index :
                Button14.configure(fg='white', bg=color)
                rgg_usb_flg.val = 1 # Rugged USB
            if "DVD" in d_index or "Blu-ray" in d_index :
                Button15.configure(fg='white', bg=color)
                odd_flg.val = 1 # ODD
            if "dGPU" in d_index :
                Button16.configure(fg='white', bg=color)
                dgpu_flg.val = 1 # dGPU
            if "Smart Card Reader2" in d_index :
                Button17.configure(fg='white', bg=color)
                scr2_flg.val = 1  # Smart Card Reader 2
            if "2nd SSD" in d_index :
                Button18.configure(fg='white', bg=color)
                ssd2nd_flg.val = 1 # 2nd SSD
            if "non Backlit KBD" in d_index :
                Button19.configure(fg='white', bg=color)
                nonbkl_flg.val = 1 # non Backlit KBD
            if "Backlit KBD" in d_index :
                Button20.configure(fg='white', bg=color)
                bkl_flg.val = 1 # Backlit KBD
            if "WLAN" in d_index :
                Button21.configure(fg='white', bg=color)
                wlan_flg.val = 1 # WLAN
        
        set_bg_color('green')
        
    # DataFrame�擾����
    read_excel('Matrix(EN)')
    sum = wwan_flg.val + wlan2nd_flg.val + gps_flg.val + dp_flg.val + cam_flg.val + finger_flg.val \
        + rfid_flg.val + scr_flg.val + batt2nd_flg.val + vga_flg.val + seri_flg.val + usb3_flg.val \
        + lan2nd_flg.val + rgg_usb_flg.val + odd_flg.val + dgpu_flg.val + scr2_flg.val + ssd2nd_flg.val \
        + nonbkl_flg.val + bkl_flg.val + wlan_flg.val
    
    # �����`�F�b�N�V�[�g�̂��ׂĂ̍��ڂ�"��"�������Ă��Ȃ������ꍇ�A���O���o�͂����Ȃ�
    if sum == 0:
        passed_flg.val = 0
        
    return 


# ���j���[�o�[�̐ݒ�
menubar = Menu(root)
# File Menu
filemenu = Menu(menubar, tearoff=False)
filemenu.add_command(label='Load', command=dispOQCcontents)

# Menu�{�^����ǉ�
menubar.add_cascade(label='Menu', menu=filemenu)

root.config(menu=menubar)

root.minsize(300,150)
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)
root.grid()

frame1 = ttk.Frame(root)
frame1.rowconfigure(0, weight=1)
frame1.columnconfigure(0,weight=1)

# �E�C���h�E�̃^�C�g���o�[�̕\�L������
root.title("OQC Launcher for FZ-55")

# �E�B���h�E�T�C�Y������(���̒���x�c�̒���)
root.geometry("700x550")

# -----------------------------------------
# ���{�K�v��OQC���ڂ����ׂĎ��s���ꂽ���𔻒�
# -----------------------------------------
def chk_pass(sum: int):
    """
    ���{�K�v��OQC���ڂ����ׂĎ��s���ꂽ���𔻒肷��B
    �Esum�ϐ��Ɋe�{�^���t���O�̒l�����v�������̂�������
    �E�@��{�V���A���ǂݍ��݌�A�usum=0 ���� PASSED_FLG=1�v
    �ł����OQC�������b�Z�[�W�\�����A���O�t�@�C�����o��
    
    model  : str
        �Ώ�SKU�i�Ԃ̕�������i�[����
    serial : str
        �Ώ�SKU�̃V���A����������i�[����
    
    passed_flg : int
        ���{�K�v��OQC���ڂ����ׂĎ��s���ꂽ���𔻒肷��
    
    Parameters
    ----------
    sum  :  int
        �{�^���t���O�̒l�̍��v�l
    """
    sum = wwan_flg.val + wlan2nd_flg.val + gps_flg.val + dp_flg.val + cam_flg.val + finger_flg.val \
        + rfid_flg.val + scr_flg.val + batt2nd_flg.val + vga_flg.val + seri_flg.val + usb3_flg.val \
        + lan2nd_flg.val + rgg_usb_flg.val + odd_flg.val + dgpu_flg.val + scr2_flg.val + ssd2nd_flg.val \
        + nonbkl_flg.val + bkl_flg.val + wlan_flg.val

    print("sum = {}".format(sum) )

    if sum == 0 and passed_flg.val == 1:
        passed_flg.val = 0
        dt_now: str = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        messagebox.showinfo('Info', 'OQC Passed. Log appeared in log.txt')
        txt.configure(state='normal')
        txt.insert('end',"OQC Passed!!   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        txt_msg: str = txt.get('1.0', tk.END)
        try:
            # with open(r'\\172.24.3.15\Imaging\QALOG\HCS_Config\OQC_Logs\log{0}_{1}_{2}.txt'.format(model.sku, model.serial, dt_now),encoding='utf-8',mode='w') as fp:
            with open(r'\\132.182.76.44\Imaging\QALOG\HCS_Config\OQC_Logs\log{0}_{1}_{2}.txt'.format(model.sku,model.serial, dt_now),encoding='utf-8',mode='w') as fp:
                fp.write(txt_msg)
                # �����t�@�C�����c���Ă�����폜����
                subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\sku.txt', shell=True)
                subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\serial.txt', shell=True)
                subprocess.run(r'fld_accs\existdel.wsf -DEL=.\exe\eco_num.txt', shell=True)
                subprocess.run(r'fld_accs\existdel.wsf -DEL=.\output\aim_data.csv', shell=True)
        except FileNotFoundError:
            messagebox.showerror('Error', 'Cannot save log file. Please ask the manager.')
            sys.exit(1)
    else:
        pass


# -----------------------
#  �e�{�^���̏������`
# -----------------------
# TODO:�{�^���̐��ɉ����Ċ֐��ǉ�or�폜���K�v
def funcWWAN(event):
    """
    �uWWAN�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if wwan_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Check WWAN module part number with device manager')

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button1.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button1.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"WWAN   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        wwan_flg.val = 0
        chk_pass(sum.val)
        return "break"


def func2ndWLAN(event):
    """
    �u2nd WLAN�v�{�^����������Ă��������Ȃ�
    """
    pass

def funcGPS(event):
    """
    �uGPS�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if gps_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button3.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button3.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"GPS   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        gps_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcDualPass(event):
    """
    �uDual Pass�v�{�^����������Ă��������Ȃ�
    """
    pass
    
def funcIRCamera(event):
    """
    �uIR Camera�v�{�^���������ꂽ��o�b�`���s + ���b�Z�[�W�\�� + ���O�ǋL
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if cam_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Test tool start. Carry out the inspection according to AIM')

    # ���s����o�b�`�̃p�X���i�[
    path1 = str( "..\\" + cam_folder)
    path2 = str( "..\\" + cam_folder + "\\" + cam_name)
    # �o�b�`�����s����
    subprocess.run(path2, cwd=path1, stdout=subprocess.PIPE, shell=False)

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button5.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button5.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"IR Camera   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        cam_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcFingerPrint(event):
    """
    �uFinger Print�v�{�^���������ꂽ��o�b�`���s + ���b�Z�[�W�\�� + ���O�ǋL
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if finger_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Test tool start. Carry out the inspection according to AIM')

    # ���s����o�b�`�̃p�X���i�[
    path1 = str( "..\\" + finger_folder)
    path2 = str( "..\\" + finger_folder + "\\" + finger_name)
    # �o�b�`�����s����
    subprocess.run(path2, cwd=path1, stdout=subprocess.PIPE, shell=False)
    
    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button6.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ� 
        Button6.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"Finger Print   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        finger_flg.val = 0
        chk_pass(sum.val)
        return "break"
    
def funcRFID(event):
    """
    �uNFC RFID�v�{�^���������ꂽ��o�b�`���s + ���b�Z�[�W�\�� + ���O�ǋL
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if rfid_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Test tool start. Carry out the inspection according to AIM')

    # ���s����o�b�`�̃p�X���i�[
    path1 = str( "..\\" + rfid_folder)
    path2 = str( "..\\" + rfid_folder + "\\" + rfid_name)
    # �o�b�`�����s����
    subprocess.run(path2, cwd=path1, stdout=subprocess.PIPE, shell=False)

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button7.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button7.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"NFC RFID   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        rfid_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcSCR(event):
    """
    �uSmart Card Reader�v�{�^���������ꂽ��o�b�`���s + ���b�Z�[�W�\�� + ���O�ǋL
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if scr_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Test tool start. Carry out the inspection according to AIM')

    # ���s����o�b�`�̃p�X���i�[
    path1 = str( "..\\" + scr_folder)
    path2 = str( "..\\" + scr_folder + "\\" + scr_name)
    # �o�b�`�����s����
    subprocess.run(path2, cwd=path1, stdout=subprocess.PIPE, shell=False)

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button8.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button8.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"Smart Card Reader   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        scr_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcSCR2(event):
    """
    �uSmart Card Reader2�v�{�^���������ꂽ��o�b�`���s + ���b�Z�[�W�\�� + ���O�ǋL
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if scr2_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Test tool start. Carry out the inspection according to AIM')

    # ���s����o�b�`�̃p�X���i�[
    path1 = str( "..\\" + scr2_folder)
    path2 = str( "..\\" + scr2_folder + "\\" + scr2_name)
    # �o�b�`�����s����
    subprocess.run(path2, cwd=path1, stdout=subprocess.PIPE, shell=False)

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button17.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button17.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"Smart Card Reader2   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        scr2_flg.val = 0
        chk_pass(sum.val)
        return "break"

def func2ndBatt(event):
    """
    �u2nd Battery�v�{�^����������Ă��������Ȃ�
    """
    pass

def funcVGA(event):
    """
    �uVGA�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if vga_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button10.configure(fg='white', bg='red')
        return "break"
    else: #Test��Pass�����Ƃ�
        Button10.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"VGA   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        vga_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcSerial(event):
    """
    �uSerial�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if seri_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button11.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button11.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"Serial   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        seri_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcUSB3(event):
    """
    �uUSB 3.0�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if usb3_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')
    
    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button12.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button12.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"USB3.0   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        usb3_flg.val = 0
        chk_pass(sum.val)
        return "break"


def func2ndLAN(event):
    """
    �u2nd LAN�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if lan2nd_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')
    
    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button13.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button13.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"2nd LAN   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        lan2nd_flg.val = 0
        chk_pass(sum.val)
        return "break"


def funcRuggedUSB(event):
    """
    �uRugged USB�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if rgg_usb_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')
    
    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button14.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button14.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"Rugged USB   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        rgg_usb_flg.val = 0
        chk_pass(sum.val)
        return "break"


def funcODD(event):
    """
    �uODD�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if odd_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button15.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button15.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"ODD   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        odd_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcdGPU(event):
    """
    �udGPU�v�{�^���������ꂽ��o�b�`���s + ���b�Z�[�W�\�� + ���O�ǋL
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if dgpu_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Test tool start. Carry out the inspection according to AIM')

    # ���s����o�b�`�̃p�X���i�[
    path1 = str( "..\\" + dgpu_folder)
    path2 = str( "..\\" + dgpu_folder + "\\" + dgpu_name)
    # �o�b�`�����s����
    subprocess.run(path2, cwd=path1, stdout=subprocess.PIPE, shell=False)

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button16.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button16.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"dGPU   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        dgpu_flg.val= 0
        chk_pass(sum.val)
        return "break"

def func2ndSSD(event):
    """
    �u2nd SSD�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if ssd2nd_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button18.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button18.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"2nd SSD   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        ssd2nd_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcNonBKL(event):
    """
    �unon Backlit KBD�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if nonbkl_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')

    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button19.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button19.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"non Backlit KBD   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        nonbkl_flg.val= 0
        chk_pass(sum.val)
        return "break"

def funcBKL(event):
    """
    �uBacklit KBD�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if bkl_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')
    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button20.configure(fg='white', bg='red')
        return "break"
    else: # Test��Pass�����Ƃ�
        Button20.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"Backlit KBD   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        bkl_flg.val = 0
        chk_pass(sum.val)
        return "break"

def funcWLAN(event):
    """
    �uWLAN�v�{�^���������ꂽ�烁�b�Z�[�W�\�� + ���O�ǋL(�o�b�`�Ȃ�)
    """
    # �����t���O�������Ă��Ȃ���΃{�^���������Ă����s���Ȃ��悤�ɂ���
    if wlan_flg.val == 0:
        return
    messagebox.showinfo('Info', 'Carry out the inspection according to AIM')
    # ���b�Z�[�W�{�b�N�X�i�͂��E�������j 
    ret = messagebox.askquestion('Confirm', 'Test Passed ?')
    if ret == 'no': # Test��fail�����Ƃ�
        messagebox.showerror('Error', 'Test Failed. Please ask the manager.')
        Button21.configure(fg='white', bg='red')
        return "break"      
    else: # Test��Pass�����Ƃ�
        Button21.configure(fg='white', bg='blue')
        dt_now = datetime.datetime.now().strftime("%X")
        txt.configure(state='normal')
        txt.insert('end',"WLAN   ")
        txt.insert('end', str(dt_now) + "\n" )
        txt.configure(state='disabled')
        wlan_flg.val = 0
        chk_pass(sum.val)
        return "break"


# ----------------------
# �e�L�X�g�{�b�N�X�̐���
# ----------------------
txt = ScrolledText(root)
txt.configure(state='disabled')
txt.place(x=450,y=200,width=200,height=300)


# -----------
# Button�ݒu
# -----------
# ��1���
Button1 = tk.Button(text='WWAN', width=20)
Button1.bind("<Button-1>", funcWWAN) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button1.place(x=50,y=50)

Button2 = tk.Button(text='2nd WLAN', width=20)
Button2.bind("<Button-1>", func2ndWLAN) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button2.place(x=50,y=100)

Button3 = tk.Button(text='GPS', width=20)
Button3.bind("<Button-1>", funcGPS) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button3.place(x=50,y=150)

Button4 = tk.Button(text='Dual Pass', width=20)
Button4.bind("<Button-1>", funcDualPass) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button4.place(x=50,y=200)

Button5 = tk.Button(text='IR Camera', width=20)
Button5.bind("<Button-1>", funcIRCamera) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button5.place(x=50,y=250)

Button6 = tk.Button(text='Finger Print', width=20)
Button6.bind("<Button-1>", funcFingerPrint) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button6.place(x=50,y=300)

Button7 = tk.Button(text='NFC RFID', width=20)
Button7.bind("<Button-1>", funcRFID) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button7.place(x=50,y=350)

Button8 = tk.Button(text='Smart Card Reader', width=20)
Button8.bind("<Button-1>", funcSCR) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button8.place(x=50,y=400)

Button17 = tk.Button(text='Smart Card Reader2', width=20)
Button17.bind("<Button-1>", funcSCR2) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button17.place(x=50,y=450)

# ��2���
Button9 = tk.Button(text='2nd Battery', width=20)
Button9.bind("<Button-1>", func2ndBatt) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button9.place(x=250,y=50)

Button10 = tk.Button(text='VGA', width=20)
Button10.bind("<Button-1>", funcVGA) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button10.place(x=250,y=100)

Button11 = tk.Button(text='Serial', width=20)
Button11.bind("<Button-1>", funcSerial) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button11.place(x=250,y=150)

Button12 = tk.Button(text='USB3.0', width=20)
Button12.bind("<Button-1>", funcUSB3) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button12.place(x=250,y=200)

Button13 = tk.Button(text='2nd LAN', width=20)
Button13.bind("<Button-1>", func2ndLAN) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button13.place(x=250,y=250)

Button14 = tk.Button(text='Rugged USB', width=20)
Button14.bind("<Button-1>", funcRuggedUSB) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button14.place(x=250,y=300)

Button15 = tk.Button(text='ODD', width=20)
Button15.bind("<Button-1>", funcODD) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button15.place(x=250,y=350)

Button16 = tk.Button(text='dGPU', width=20)
Button16.bind("<Button-1>", funcdGPU) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button16.place(x=250,y=400)

Button18 = tk.Button(text='2nd SSD', width=20)
Button18.bind("<Button-1>", func2ndSSD) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button18.place(x=250,y=450)

# ��3���
Button19 = tk.Button(text='non Backlit KBD', width=20)
Button19.bind("<Button-1>", funcNonBKL) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button19.place(x=450,y=50)

Button20 = tk.Button(text='Backlit KBD', width=20)
Button20.bind("<Button-1>", funcBKL) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button20.place(x=450,y=100)

Button21 = tk.Button(text='WLAN', width=20)
Button21.bind("<Button-1>", funcWLAN) # �{�^���������ꂽ�Ƃ��Ɏ��s�����֐����o�C���h����
Button21.place(x=450,y=150)


#�\�t�g�E�F�A�̎��s���e�̏���(�I��)
root.mainloop()
