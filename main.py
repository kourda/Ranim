from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5 import QtGui
from threading import *
from twilio.rest import Client
import os
import yagmail
import sys
from time import *
from xlrd import open_workbook
import pickle


class Ui(QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        self.email = "raniimmny@gmail.com"
        self.password = "sqxmequuvvilqtdo"
        self.reemail= "ranimmanai123@gmail.com"
        self.sid = 'AC73511681deeaa375cd3a88cf7430580a'
        self.auth = 'be4403aa9514517a9a8e0ef5d3f5c2ba'
        self.hostnum = '+13517776598'
        self.reciver = '+21629726238'
        self.wb1 = open_workbook('stats2G.xls').sheet_by_index(0)
        self.wb2 = open_workbook('stats3G.xls').sheet_by_index(0)
        self.wb3 = open_workbook('stats4G.xls').sheet_by_index(0)
        self.setWindowIcon(QtGui.QIcon('icon.ico'))
        self.makeDataFile()
        self.makeDirs()
        self.message = self.getData()
        self.sms = []
        self.save_info2g(self.wb1)
        self.save_info3g(self.wb2)
        self.save_info4g(self.wb3)
        self.thread()
        uic.loadUi('init.ui', self)
        self.show()
        self.update()
        self.update1.clicked.connect(self.changeData2G)
        self.update2.clicked.connect(self.changeData3G)
        self.update3.clicked.connect(self.changeData4G)
        # TCH 2g
        self.tch2g_1.clicked.connect(self.tch2g_1_stats)
        self.tch2g_2.clicked.connect(self.tch2g_2_stats)
        self.tch2g_3.clicked.connect(self.tch2g_3_stats)
        # Disp 2g
        self.dis2g_1.clicked.connect(self.dis2g_1_stats)
        self.dis2g_2.clicked.connect(self.dis2g_2_stats)
        self.dis2g_3.clicked.connect(self.dis2g_3_stats)
        # CSDROP 2g
        self.calldrop2g_1.clicked.connect(self.calldrop2g_1_stats)
        self.calldrop2g_2.clicked.connect(self.calldrop2g_2_stats)
        self.calldrop2g_3.clicked.connect(self.calldrop2g_3_stats)

        # DISP 3g
        self.t1.clicked.connect(self.t1_stats)
        self.t2.clicked.connect(self.t2_stats)
        self.t3.clicked.connect(self.t3_stats)
        
        # CSDROP 3g
        self.q1.clicked.connect(self.q1_stats)
        self.q2.clicked.connect(self.q2_stats)
        self.q3.clicked.connect(self.q3_stats)

        # PSDROP 3g
        self.w1.clicked.connect(self.w1_stats)
        self.w2.clicked.connect(self.w2_stats)
        self.w3.clicked.connect(self.w3_stats)

        # CSSR CS 3g
        self.e1.clicked.connect(self.e1_stats)
        self.e2.clicked.connect(self.e2_stats)
        self.e3.clicked.connect(self.e3_stats)

        # CSSR PS 3g
        self.y1.clicked.connect(self.y1_stats)
        self.y2.clicked.connect(self.y2_stats)
        self.y3.clicked.connect(self.y3_stats)

        # Disponibilite 4g
        self.u1.clicked.connect(self.u1_stats)
        self.u2.clicked.connect(self.u2_stats)
        self.u3.clicked.connect(self.u3_stats)

        # Calldrop 4g
        self.i1.clicked.connect(self.i1_stats)
        self.i2.clicked.connect(self.i2_stats)
        self.i3.clicked.connect(self.i3_stats)
        
        # SSSR 4G
        self.o1.clicked.connect(self.o1_stats)
        self.o2.clicked.connect(self.o2_stats)
        self.o3.clicked.connect(self.o3_stats)
    def sendmessage(self):
        client = Client(self.sid, self.auth)
        for i in range(1):
            message = client.messages.create(body=self.sms[i],from_=self.hostnum,to=self.reciver)
        print('done')
    def send_mineur_email(self):
        yag = yagmail.SMTP(self.email, self.password)
        # 2G
        msg ="Réseau: 2G \n"
        file1 = open(os.path.join('2g','tch','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file1)
                msg+= f"kpi: TCH - Mineur - { e['tch'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file1.close()

        file2 = open(os.path.join('2g','disp','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: Disponibilite - Mineur - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file2.close()

        file3 = open(os.path.join('2g','calldrop','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: Calldrop - Mineur - { e['calldrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file3.close()

        yag.send(self.reemail,'2G-Mineur',msg)

        msg ="Réseau: 3G \n"
        file1 = open(os.path.join('3g','disp','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file1)
                msg+= f"kpi: Dipsonibilite - Mineur - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file1.close()

        file2 = open(os.path.join('3g','csdrop','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: CsDrop - Mineur - { e['csdrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file2.close()

        file3 = open(os.path.join('3g','psdrop','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: PsDrop - Mineur - { e['psdrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file3.close()

        file4 = open(os.path.join('3g','cssrcs','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file4)
                msg+= f"kpi: Cssrcs - Mineur - { e['cssrcs'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file4.close()
        file5 = open(os.path.join('3g','cssrps','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file5)
                msg+= f"kpi: CssrPs - Mineur - { e['cssrps'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file5.close()
        yag.send(self.reemail,'3G-Mineur',msg)

        msg ="Réseau: 4G \n"
        file1 = open(os.path.join('4g','disp','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file1)
                msg+= f"kpi: DISP - Mineur - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file1.close()

        file2 = open(os.path.join('4g','calldrop','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: CallDrop - Mineur - { e['calldrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file2.close()

        file3 = open(os.path.join('4g','sssr','mineur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: SSSR - Mineur - { e['sssr'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file3.close()
        yag.send(self.reemail,'4G-Mineur',msg)
    
    def send_majeur_email(self):
        yag = yagmail.SMTP(self.email, self.password)
        # 2G
        msg ="Réseau: 2G \n"
        file1 = open(os.path.join('2g','tch','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file1)
                msg+= f"kpi: TCH - Majeur - { e['tch'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file1.close()

        file2 = open(os.path.join('2g','disp','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: Disponibilite - Majeur - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file2.close()

        file3 = open(os.path.join('2g','calldrop','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: Calldrop - Majeur - { e['calldrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file3.close()

        yag.send(self.reemail,'2G-Majeur',msg)

        msg ="Réseau: 3G \n"
        file1 = open(os.path.join('3g','disp','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file1)
                msg+= f"kpi: Dipsonibilite - Majeur - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file1.close()

        file2 = open(os.path.join('3g','csdrop','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: CsDrop - Majeur - { e['csdrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file2.close()

        file3 = open(os.path.join('3g','psdrop','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: PsDrop - Majeur - { e['psdrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file3.close()

        file4 = open(os.path.join('3g','cssrcs','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file4)
                msg+= f"kpi: Cssrcs - Majeur - { e['cssrcs'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file4.close()
        file5 = open(os.path.join('3g','cssrps','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file5)
                msg+= f"kpi: CssrPs - Majeur - { e['cssrps'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file5.close()
        yag.send(self.reemail,'3G-Majeur',msg)

        msg ="Réseau: 4G \n"
        file1 = open(os.path.join('4g','disp','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file1)
                msg+= f"kpi: DISP - Critique - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file1.close()

        file2 = open(os.path.join('4g','calldrop','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: CallDrop - Majeur - { e['calldrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file2.close()

        file3 = open(os.path.join('4g','sssr','majeur.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: SSSR - Majeur - { e['sssr'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file3.close()
        yag.send(self.reemail,'4G-Majeur',msg)
    

    def send_critique_email(self):
        client = Client(self.sid, self.auth)
        yag = yagmail.SMTP(self.email, self.password)
        # 2G
        msg ="Réseau: 2G \n"
        file1 = open(os.path.join('2g','tch','critique.dat'),'rb')
        while True:
            try:
                # si mms gateaway provided by telecom
                e = pickle.load(file1)
                msg+= f"kpi: TCH - Critique - { e['tch'] } | Id: {e['id']} | Nom: {e['nom']} \n"
                #message = client.messages.create(body=f"2G : kpi: TCH - Critique - { e['tch'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
            except:
                break
        message = client.messages.create(body=f"2G : kpi: TCH - Critique - { e['tch'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
        file1.close()

        file2 = open(os.path.join('2g','disp','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: Disponibilite - Critique - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
                #message = client.messages.create(body=f"2G : kpi: Disponibilite - Critique - { e['tch'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
            except:
                break
        file2.close()
        message = client.messages.create(body=f"2G : kpi: Disponibilite - Critique - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
        file3 = open(os.path.join('2g','calldrop','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: Calldrop - Critique - { e['calldrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
                #message = client.messages.create(body=f"2G : kpi: Calldrop - Critique - { e['calldrop'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
            except:
                break
        file3.close()

        yag.send(self.reemail,'2G-Critique',msg)

        msg ="Réseau: 3G \n"
        file1 = open(os.path.join('3g','disp','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file1)
                msg+= f"kpi: Dipsonibilite - Critique - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        message = client.messages.create(body=f"3G : kpi: Disponibilite - Critique - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
        file1.close()

        file2 = open(os.path.join('3g','csdrop','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: CsDrop - Critique - { e['csdrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        message = client.messages.create(body=f"3G : kpi: Csdrop - Critique - { e['csdrop'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
        file2.close()

        file3 = open(os.path.join('3g','psdrop','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: PsDrop - Critique - { e['psdrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file3.close()

        file4 = open(os.path.join('3g','cssrcs','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file4)
                msg+= f"kpi: Cssrcs - Critique - { e['cssrcs'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file4.close()
        file5 = open(os.path.join('3g','cssrps','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file5)
                msg+= f"kpi: CssrPs - Critique - { e['cssrps'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        message = client.messages.create(body=f"3G : kpi: Cssrps - Critique - { e['cssrps'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
        file5.close()
        yag.send(self.reemail,'3G-Critique',msg)

        msg ="Réseau: 4G \n"
        file1 = open(os.path.join('4g','disp','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file1)
                msg+= f"kpi: DISP - Critique - { e['disp'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        file1.close()

        file2 = open(os.path.join('4g','calldrop','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file2)
                msg+= f"kpi: CallDrop - Critique - { e['calldrop'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        message = client.messages.create(body=f"4G : kpi: CallDrop - Critique - { e['calldrop'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
        file2.close()

        file3 = open(os.path.join('4g','sssr','critique.dat'),'rb')
        while True:
            try:
                e = pickle.load(file3)
                msg+= f"kpi: SSSR - Critique - { e['sssr'] } | Id: {e['id']} | Nom: {e['nom']} \n"
            except:
                break
        message = client.messages.create(body=f"4G : kpi: SSSR - Critique - { e['sssr'] } | Id: {e['id']} | Nom: {e['nom']}",from_=self.hostnum,to=self.reciver)
        file3.close()
        yag.send(self.reemail,'4G-Critique',msg)

    def o1_stats(self):
        app2 = Ui2(self,os.path.join('4g','sssr','critique.dat'),'sssr')
        app2.exec_()

    def o2_stats(self):
        app2 = Ui2(self,os.path.join('4g','sssr','majeur.dat'),'sssr')
        app2.exec_()

    def o3_stats(self):
        app2 = Ui2(self,os.path.join('4g','sssr','mineur.dat'),'sssr')
        app2.exec_()


    def i1_stats(self):
        app2 = Ui2(self,os.path.join('4g','calldrop','critique.dat'),'calldrop')
        app2.exec_()

    def i2_stats(self):
        app2 = Ui2(self,os.path.join('4g','calldrop','majeur.dat'),'calldrop')
        app2.exec_()

    def i3_stats(self):
        app2 = Ui2(self,os.path.join('4g','calldrop','mineur.dat'),'calldrop')
        app2.exec_()

    def u1_stats(self):
        app2 = Ui2(self,os.path.join('4g','disp','critique.dat'),'disp')
        app2.exec_()

    def u2_stats(self):
        app2 = Ui2(self,os.path.join('4g','disp','majeur.dat'),'disp')
        app2.exec_()

    def u3_stats(self):
        app2 = Ui2(self,os.path.join('4g','disp','mineur.dat'),'disp')
        app2.exec_()


    def y1_stats(self):
        app2 = Ui2(self,os.path.join('3g','cssrps','critique.dat'),'cssrps')
        app2.exec_()

    def y2_stats(self):
        app2 = Ui2(self,os.path.join('3g','cssrps','majeur.dat'),'cssrps')
        app2.exec_()

    def y3_stats(self):
        app2 = Ui2(self,os.path.join('3g','cssrps','mineur.dat'),'cssrps')
        app2.exec_()

    def e1_stats(self):
        app2 = Ui2(self,os.path.join('3g','cssrcs','critique.dat'),'cssrcs')
        app2.exec_()

    def e2_stats(self):
        app2 = Ui2(self,os.path.join('3g','cssrcs','majeur.dat'),'cssrcs')
        app2.exec_()

    def e3_stats(self):
        app2 = Ui2(self,os.path.join('3g','cssrcs','mineur.dat'),'cssrcs')
        app2.exec_()

    def w1_stats(self):
        app2 = Ui2(self,os.path.join('3g','psdrop','critique.dat'),'psdrop')
        app2.exec_()

    def w2_stats(self):
        app2 = Ui2(self,os.path.join('3g','psdrop','majeur.dat'),'psdrop')
        app2.exec_()

    def w3_stats(self):
        app2 = Ui2(self,os.path.join('3g','psdrop','mineur.dat'),'psdrop')
        app2.exec_()


    def q1_stats(self):
        app2 = Ui2(self,os.path.join('3g','csdrop','critique.dat'),'csdrop')
        app2.exec_()

    def q2_stats(self):
        app2 = Ui2(self,os.path.join('3g','csdrop','majeur.dat'),'csdrop')
        app2.exec_()

    def q3_stats(self):
        app2 = Ui2(self,os.path.join('3g','csdrop','mineur.dat'),'csdrop')
        app2.exec_()

    def t1_stats(self):
        app2 = Ui2(self,os.path.join('3g','disp','critique.dat'),'disp')
        app2.exec_()

    def t2_stats(self):
        app2 = Ui2(self,os.path.join('3g','disp','majeur.dat'),'disp')
        app2.exec_()

    def t3_stats(self):
        app2 = Ui2(self,os.path.join('3g','disp','mineur.dat'),'disp')
        app2.exec_()

    def calldrop2g_1_stats(self):
        app2 = Ui2(self,os.path.join('2g','calldrop','critique.dat'),'calldrop')
        app2.exec_()

    def calldrop2g_2_stats(self):
        app2 = Ui2(self,os.path.join('2g','calldrop','majeur.dat'),'calldrop')
        app2.exec_()

    def calldrop2g_3_stats(self):
        app2 = Ui2(self,os.path.join('2g','calldrop','mineur.dat'),'calldrop')
        app2.exec_()

    def dis2g_1_stats(self):
        app2 = Ui2(self,os.path.join('2g','disp','critique.dat'),'disp')
        app2.exec_()

    def dis2g_2_stats(self):
        app2 = Ui2(self,os.path.join('2g','disp','majeur.dat'),'disp')
        app2.exec_()

    def dis2g_3_stats(self):
        app2 = Ui2(self,os.path.join('2g','disp','mineur.dat'),'disp')
        app2.exec_()

    def tch2g_1_stats(self):
        app2 = Ui2(self,os.path.join('2g','tch','critique.dat'),'tch')
        app2.exec_()

    def tch2g_2_stats(self):
        app2 = Ui2(self,os.path.join('2g','tch','majeur.dat'),'tch')
        app2.exec_()

    def tch2g_3_stats(self):
        app2 = Ui2(self,os.path.join('2g','tch','mineur.dat'),'tch')
        app2.exec_()

    def thread(self):
        t1 = Thread(target=self.sendEmails)
        t1.start()

    def sendEmails(self):
        self.sendmessage()
        self.send_critique_email()
        print('done1')
        self.send_majeur_email()
        print('done2')
        self.send_mineur_email()
        print('done3')
        while True:
            sleep(3200)
            self.sms=[]
            self.save_info2g(self.wb1)
            self.save_info3g(self.wb2)
            self.save_info4g(self.wb3)
            self.sendmessage()
            self.send_critique_email()
            print('done1')
            self.send_majeur_email()
            print('done2')
            self.send_mineur_email()
            print('done3')

    def makeDirs(self):
        for i in ['2g', '3g', '4g']:
            try:
                os.makedirs(i)
            except:
                pass
        for j in ['tch','disp','calldrop']:
            try:
                os.makedirs(os.path.join('2g',j))
            except:
                pass
        for k in ['disp','csdrop','psdrop','cssrcs','cssrps']:
            try:
                os.makedirs(os.path.join('3g',k))
            except:
                pass
        for d in ['disp','calldrop','sssr']:
            try:
                os.makedirs(os.path.join('4g',d))
            except:
                pass
    def save_info2g(self,sheet):
        # TCH
        file1 = open(os.path.join('2g','tch','mineur.dat'),'wb')
        file2 = open(os.path.join('2g','tch','majeur.dat'),'wb')
        file3 = open(os.path.join('2g','tch','critique.dat'),'wb')
        # DISP
        file4 = open(os.path.join('2g','disp','mineur.dat'),'wb')
        file5 = open(os.path.join('2g','disp','majeur.dat'),'wb')
        file6 = open(os.path.join('2g','disp','critique.dat'),'wb')
        # Call drop
        file7 = open(os.path.join('2g','calldrop','mineur.dat'),'wb')
        file8 = open(os.path.join('2g','calldrop','majeur.dat'),'wb')
        file9 = open(os.path.join('2g','calldrop','critique.dat'),'wb')
        for i in range(sheet.nrows-1):
            e = {}
            e['id'] = sheet.cell_value(i+1, 1)
            e['nom'] = sheet.cell_value(i+1, 2)
            e['hour'] = sheet.cell_value(i+1,3)
            e['date'] = sheet.cell_value(i+1,4)
            e['t1'] = sheet.cell_value(i+1,5)
            e['calldrop'] = sheet.cell_value(i+1, 6)
            if type(e['calldrop']) == str:
                e['calldrop'] = 0
            e['disp'] = sheet.cell_value(i+1, 7)
            if type(e['disp']) == str:
                e['disp'] = 0
            e['tch'] = sheet.cell_value(i+1, 8)
            if type(e['tch']) == str:
                e['tch'] = 0

            
            if e['tch'] < self.message['2g']['tch'][0]:
                pickle.dump(e,file1)
            elif e['tch'] < self.message['2g']['tch'][1]:
                pickle.dump(e,file2)
            else:
                pickle.dump(e,file3)

                
            if e['disp'] > self.message['2g']['disp'][0]:
                pickle.dump(e,file4)
            elif e['disp'] > self.message['2g']['disp'][1]:
                pickle.dump(e,file5)
            else:
                pickle.dump(e,file6)
            
            if e['calldrop'] < self.message['2g']['calldrop'][0]:
                pickle.dump(e,file7)
            elif e['calldrop'] <  self.message['2g']['calldrop'][1]:
                pickle.dump(e,file8)
            else:
                pickle.dump(e,file9)
            if (e['disp']==100 and e['t1']==0):
                self.sms.append(f"2G: Nécessite un RESET du site  {e['nom']} {e['id']}")
                

    def save_info3g(self,sheet):
        file1 = open(os.path.join('3g','disp','mineur.dat'),'wb')
        file2 = open(os.path.join('3g','disp','majeur.dat'),'wb')
        file3 = open(os.path.join('3g','disp','critique.dat'),'wb')
    
        file4 = open(os.path.join('3g','csdrop','mineur.dat'),'wb')
        file5 = open(os.path.join('3g','csdrop','majeur.dat'),'wb')
        file6 = open(os.path.join('3g','csdrop','critique.dat'),'wb')
    
        file7 = open(os.path.join('3g','psdrop','mineur.dat'),'wb')
        file8 = open(os.path.join('3g','psdrop','majeur.dat'),'wb')
        file9 = open(os.path.join('3g','psdrop','critique.dat'),'wb')
    
        file10 = open(os.path.join('3g','cssrcs','mineur.dat'),'wb')
        file11 = open(os.path.join('3g','cssrcs','majeur.dat'),'wb')
        file12 = open(os.path.join('3g','cssrcs','critique.dat'),'wb')
    
        file13 = open(os.path.join('3g','cssrps','mineur.dat'),'wb')
        file14 = open(os.path.join('3g','cssrps','majeur.dat'),'wb')
        file15 = open(os.path.join('3g','cssrps','critique.dat'),'wb')
    
        sheet = open_workbook('stats3G.xls').sheet_by_index(0)
        for i in range(sheet.nrows-1):
            e = {}
            e['id'] = sheet.cell_value(i+1, 3)
            e['nom'] = sheet.cell_value(i+1, 2)
            e['disp'] = sheet.cell_value(i+1, 6)
            e['t1'] = sheet.cell_value(i+1,5)
            e['t2'] = sheet.cell_value(i+1,4)
            if type(e['disp']) == str:
                e['disp'] = 0
            e['csdrop'] = sheet.cell_value(i+1, 7)
            if type(e['csdrop']) == str:
                e['csdrop'] = 0
            e['psdrop'] = sheet.cell_value(i+1, 8)
            if type(e['psdrop']) == str:
                e['psdrop'] = 0
            e['cssrcs'] = sheet.cell_value(i+1, 9)
            if type(e['cssrcs']) == str:
                e['cssrcs'] = 0
            e['cssrps'] = sheet.cell_value(i+1, 10)
            if type(e['cssrps']) == str:
                e['cssrps'] = 0
    
            if e['disp'] >= self.message['3g']['disp'][0]:
                pickle.dump(e,file1)
            elif e['disp'] >= self.message['3g']['disp'][1]:
                pickle.dump(e,file2)
            else:
                pickle.dump(e,file3)
            if e['csdrop'] < self.message['3g']['csdrop'][0]:
                pickle.dump(e,file4)
            elif e['csdrop'] < self.message['3g']['csdrop'][1]:
                pickle.dump(e,file5)
            else:
                pickle.dump(e,file6)
            
            if e['psdrop'] < self.message['3g']['psdrop'][0]:
                pickle.dump(e,file7)
            elif e['psdrop'] < self.message['3g']['psdrop'][1]:
                pickle.dump(e,file8)
            else:
                pickle.dump(e,file9)
    
            if e['cssrcs'] > self.message['3g']['cssrcs'][0]:
                pickle.dump(e,file10)
            elif e['cssrcs'] > self.message['3g']['cssrcs'][1]:
                pickle.dump(e,file11)
            else:
                pickle.dump(e,file12)
                
            if e['cssrps'] > self.message['3g']['cssrps'][0]:
                pickle.dump(e,file13)
            elif e['cssrps'] > self.message['3g']['cssrps'][1]:
                pickle.dump(e,file14)
            else:
                pickle.dump(e,file15)
                
            if int(e['disp'])==100 and (e['t1']==0 or e['t2']==0):
                self.sms.append(f"3G: Nécessite un RESET du site  {e['nom']} {e['id']}")

    def save_info4g(self,sheet):
        # DISP
        file1 = open(os.path.join('4g','disp','mineur.dat'),'wb')
        file2 = open(os.path.join('4g','disp','majeur.dat'),'wb')
        file3 = open(os.path.join('4g','disp','critique.dat'),'wb')
        # CALLDROP
        file4 = open(os.path.join('4g','calldrop','mineur.dat'),'wb')
        file5 = open(os.path.join('4g','calldrop','majeur.dat'),'wb')
        file6 = open(os.path.join('4g','calldrop','critique.dat'),'wb')
        # SSSR
        file7 = open(os.path.join('4g','sssr','mineur.dat'),'wb')
        file8 = open(os.path.join('4g','sssr','majeur.dat'),'wb')
        file9 = open(os.path.join('4g','sssr','critique.dat'),'wb')
        for i in range(sheet.nrows-1):
            e = {}
            e['id'] = sheet.cell_value(i+1, 1)
            e['nom'] = sheet.cell_value(i+1, 0)
            e['disp'] = sheet.cell_value(i+1, 7)
            if type(e['disp']) == str:
                e['disp'] = 0
            e['traffic'] = sheet.cell_value(i+1, 6)
            if type(e['traffic']) == str:
                e['traffic'] = 0
            e['sssr'] = sheet.cell_value(i+1, 5)
            if type(e['sssr']) == str:
                e['sssr'] = 0
            e['calldrop'] = sheet.cell_value(i+1, 4)
            if type(e['calldrop']) == str:
                e['calldrop'] = 0

            if e['disp'] > self.message['4g']['disp'][0]:
                pickle.dump(e,file1)
            elif e['disp'] > self.message['4g']['disp'][1]:
                pickle.dump(e,file2)
            else:
                pickle.dump(e,file3)

            if e['calldrop'] < self.message['4g']['calldrop'][0]:
                pickle.dump(e,file4)
            elif e['calldrop'] < self.message['4g']['calldrop'][1]:
                pickle.dump(e,file5)
            else:
                pickle.dump(e,file6)

            if e['sssr'] > self.message['4g']['sssr'][0]:
                pickle.dump(e,file7)
            elif e['sssr'] > self.message['4g']['sssr'][1]:
                pickle.dump(e,file8)
            else:
                pickle.dump(e,file9)
            if int(e['disp'])==100 and e['traffic']==0:
                self.sms.append(f"4G: Nécessite un RESET du site  {e['nom']} {e['id']}")
    
    def update(self):
        # 2G
        self.tch2g1.setValue(self.message['2g']['tch'][0])
        self.tch2g2.setValue(self.message['2g']['tch'][1])
        self.dis2g1.setValue(self.message['2g']['disp'][0])
        self.dis2g2.setValue(self.message['2g']['disp'][1])
        self.cdrop2g1.setValue(self.message['2g']['calldrop'][0])
        self.cdrop2g2.setValue(self.message['2g']['calldrop'][1])
        # 3G
        self.dis3g1.setValue(self.message['3g']['disp'][0])
        self.dis3g2.setValue(self.message['3g']['disp'][1])
        self.csdrop3g1.setValue(self.message['3g']['csdrop'][0])
        self.csdrop3g2.setValue(self.message['3g']['csdrop'][1])
        self.psdrop3g1.setValue(self.message['3g']['psdrop'][0])
        self.psdrop3g2.setValue(self.message['3g']['psdrop'][1])
        self.cssrcs3g1.setValue(self.message['3g']['cssrcs'][0])
        self.cssrcs3g2.setValue(self.message['3g']['cssrcs'][1])
        self.cssrps3g1.setValue(self.message['3g']['cssrps'][0])
        self.cssrps3g2.setValue(self.message['3g']['cssrps'][1])
        # 4G
        self.disp4g1.setValue(self.message['4g']['disp'][0])
        self.disp4g2.setValue(self.message['4g']['disp'][1])
        self.calldrop4g1.setValue(self.message['4g']['calldrop'][0])
        self.calldrop4g2.setValue(self.message['4g']['calldrop'][1])
        self.sssr4g1.setValue(self.message['4g']['sssr'][0])
        self.sssr4g2.setValue(self.message['4g']['sssr'][1])

    def changeData4G(self):
        self.message['4g']['disp'][0]=self.disp4g1.value()
        self.message['4g']['disp'][1]=self.disp4g2.value()

        self.message['4g']['calldrop'][0]=self.calldrop4g1.value()
        self.message['4g']['calldrop'][1]=self.calldrop4g2.value()

        self.message['4g']['sssr'][0]=self.sssr4g1.value()
        self.message['4g']['sssr'][1]=self.sssr4g2.value()

        self.changeMainMessage(self.message)
        self.save_info4g(self.wb3)
    
    def changeData3G(self):
        self.message['3g']['disp'][0]=self.dis3g1.value()
        self.message['3g']['disp'][1]=self.dis3g2.value()

        self.message['3g']['csdrop'][0]=self.csdrop3g1.value()
        self.message['3g']['csdrop'][1]=self.csdrop3g2.value()

        self.message['3g']['psdrop'][0]=self.psdrop3g1.value()
        self.message['3g']['psdrop'][1]=self.psdrop3g2.value()

        self.message['3g']['cssrcs'][0]=self.cssrcs3g1.value()
        self.message['3g']['cssrcs'][1]=self.cssrcs3g2.value()

        self.message['3g']['cssrps'][0]=self.cssrps3g1.value()
        self.message['3g']['cssrps'][1]=self.cssrps3g2.value()
        self.changeMainMessage(self.message)
        self.save_info3g(self.wb2)
    
    def changeData2G(self):
        self.message['2g']['tch'][0]=self.tch2g1.value()
        self.message['2g']['tch'][1]=self.tch2g2.value()

        self.message['2g']['disp'][0]=self.dis2g1.value()
        self.message['2g']['disp'][1]=self.dis2g2.value()

        self.message['2g']['calldrop'][0]=self.cdrop2g1.value()
        self.message['2g']['calldrop'][1]=self.cdrop2g2.value()
        self.changeMainMessage(self.message)
        self.save_info2g(self.wb1)


    def makeDataFile(self):
        if not(os.path.exists('data.dat')):
            msg = {}
            msg['2g']={
                'tch': [0.03,0.1],
                'disp': [98,90],
                'calldrop':[0.8,1.5]
            }
            msg['3g']={
                'disp': [98,90],
                'csdrop': [0.8,1.5],
                'psdrop': [2.0,5.0],
                'cssrcs': [98,95],
                'cssrps': [90,80]
            }
            msg['4g']={
                'disp': [98,90],
                'calldrop':[0.8,1.5],
                'sssr': [90,80]
            }
            self.changeMainMessage(msg)        
    def getData(self):
        with open('data.dat','rb') as fileData:
            return(pickle.load(fileData))
    def changeMainMessage(self,new):
        with open('data.dat','wb') as fileData:
            pickle.dump(new,fileData)



class Ui2(QDialog):
    def __init__(self,parent=None,*args,**kwargs):
        super().__init__(parent)
        uic.loadUi("second.ui",self)
        self.setWindowIcon(QtGui.QIcon('icon.ico'))
        self.file = args[0]
        self.sizeof = self.size()
        self.champ = args[1]
        self.t1.setRowCount(self.sizeof)
        self.t1.setColumnWidth(1,300)
        self.t1.setColumnWidth(2,300)
        self.t1.setColumnWidth(3,300)
        self.dataShow()
    def size(self):
        f1 = open(self.file,'rb')
        i=0
        while True:
            try:
                e = pickle.load(f1)
                i+=1
            except:
                break
        return i
    def dataShow(self):
        f1 = open(self.file,'rb')
        for j in range(self.sizeof):
            e = pickle.load(f1)
            self.t1.setItem(j,0,QTableWidgetItem(str(e['id'])))
            self.t1.setItem(j,1,QTableWidgetItem(str(e['nom'])))
            self.t1.setItem(j,2,QTableWidgetItem(str(e[self.champ])))
        f1.close()
app = QApplication(sys.argv)
UIWindow = Ui()
app.exec_()
