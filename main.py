from scipy.stats import linregress
from win32com.client import Dispatch
from time import sleep
from os import remove
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import pyvisa
import sys
import csv

ENABLE_MONODAQ = True

class StateMachine:

    def __init__(self):
        #print("INITIALISATION")

        self.error_messages = []
        self.rm = pyvisa.ResourceManager()
        self.keithley = None
        try:
            self.keithley = self.rm.open_resource('ASRL5::INSTR')
        except pyvisa.errors.VisaIOError:
            self.error_messages.append(["Impossible de se connecter au keithley"])
            self.state = "ERROR"
            return

        self.keithley.baud_rate = 9600
        self.keithley.timeout = 25000
        self.keithley.read_termination = '\r'
        self.keithley.write_termination = '\r'

        self.sample_thickness = None #um
        self.start_current = None # limite inférieure
        self.stop_current = None
        self.nbr_mesures = None
        self.V_compliance = 50 # V
        self.step = None

        self.mesures = []
        self.currents = []
        self.voltages = []

        self.slope = 0
        self.intercept = 0
        self.regression_line = 0

        self.pressureLimits = [4500, 5500]
        self.setupPath = r"C:\Users\33618\Desktop\cours\IUT\Projet slovénie\projet summer camp\setups dewesoftx\test 1.dxs"

        self.tempFileNbr = 0
        self.csvFiles = []
        self.sample_name = "error"
        self.file_name = "error.csv"

        self.previous_state = "INIT"
        self.state = "CHECK PRESSURE"

    def run(self):

        while self.state != "STOP":

            if self.state == "CHECK PRESSURE":
                self.checkpressure()
            elif self.state == "DEFINITION PARAMETRES":
                self.defparametres()
            elif self.state == "MEASURE":
                self.measure()
            elif self.state == "SAVE":
                self.save()
            elif self.state == "ERROR":
                self.error()
            else:
                self.error()
        self.stop()

    def checkpressure(self):
        #print("CHECK PRESSURE")

        if ENABLE_MONODAQ:
            # create DCOM object
            print("Création de l'objet DCOM.")
            dw = Dispatch("Dewesoft.App")

            # open Dewesoft
            print("Initialisation de DewesoftX ... ", end="")
            sys.stdout.flush()
            dw.Init()

            dw.Enabled = 1
            dw.Visible = 0

            # set window dimensions
            dw.Top = 0
            dw.Left = 0
            dw.Width = 1920
            dw.Height = 1080
            print('Initialisation terminée.')

            # change PATH to Setup file accordingly
            print("Chargement du fichier configuration ... ", end="")

            dw.LoadSetup(self.setupPath)

            print("Chargement terminé.")

            # build channel list
            print("Construction de la liste des canaux.")
            dw.Data.BuildChannelList()
            conn_list = [dw.Data.UsedChannels.Item(i) for i in range(dw.Data.UsedChannels.Count)]
            dw.Start()
            print("\n\n\n\n\n\n---- VERIFICATION DE LA PRESSION ----")
            consecutive = 0
            while True:
                canal = 4
                nbr = 5
                #mesure le voltage à la broche 1
                m = 0
                for j in range(nbr):
                    conn = conn_list[canal]
                    if conn.DBDataSize >= 1:  # vérifie si le buffer du canal 1 peut être lu
                        BufPos = conn.DBPos
                        PosToRead = conn.DBPos - 1
                        if BufPos == 0:
                            PosToRead = PosToRead + conn.DBDataSize
                        a = conn.DBValues(PosToRead)
                        m += a
                    sleep(0.01)
                m = m/nbr
                if m <= 0:
                    m = float("inf")

                if m > self.pressureLimits[1]:
                    print(f"pression trop faible ({m:.0f} ohms)", end="\r")
                    consecutive = 0
                elif m < self.pressureLimits[0]:
                    print(f"pression trop forte  ({m:.0f} ohms)", end="\r")
                else:
                    print(f"pression correcte    ({m:.0f} ohms)", end="\r")
                    consecutive += 1

                if consecutive >= 50:
                    dw.Stop()
                    break
        self.previous_state = self.state
        self.state = "DEFINITION PARAMETRES"

    def defparametres(self):

        print("\n\n\n\n\n\n---- DEFINITION DES PARAMETRES ----")
        print("\n\n")

        if self.previous_state == "CHECK PRESSURE":
            self.file_name = input("Nom du fichier de sauvegarde : ")
        self.sample_name = input("Nom de l'échantillon : ")
        self.sample_thickness = float(input("Epaisseur de l'échantillon (um) : "))
        self.stop_current = float(input("Courant maximum (A): "))
        self.start_current = - self.stop_current
        self.nbr_mesures = int(input("Nombre de mesures : "))
        self.step = abs(self.stop_current - self.start_current) / (self.nbr_mesures - 1)

        if input("\n\nLes paramètres sont-ils corrects ? (o/n) : ") == "o":
            self.previous_state = self.state
            self.state = "MEASURE"
        else:
            self.previous_state = self.state
            self.state = "DEFINITION DES PARAMETRES"
        print("\n\n")

    def measure(self):
        print("MEASURE")

        self.keithley.write("*RST")

        self.keithley.write(':SENS:FUNC:CONC OFF')                 # Turn off concurrent functions
        self.keithley.write(':SOUR:FUNC CURR')                     # Current source function
        if self.keithley.query('SENS:FUNC?') != "\"VOLT:DC\"":
            self.keithley.write(':SENS:FUNC:ON VOLT')              # Volts sense function
        self.keithley.write(f':SENS:VOLT:PROT {self.V_compliance}')     # 1V voltage compliance.
        self.keithley.write(f':SOUR:CURR:START {self.start_current}')   # start current
        self.keithley.write(f':SOUR:CURR:STOP {self.stop_current}')     # stop current
        self.keithley.write(f':SOUR:CURR:STEP {self.step}')             # step current
        self.keithley.write(':SOUR:CURR:MODE SWE')                 # Select current sweep mode.
        #his command should normally be sent after START, STOP, and STEP to avoid
        #delays caused by rebuilding sweep when each command is sent
        self.keithley.write(':SOUR:SWE:RANG AUTO')                 # Auto source ranging
        self.keithley.write(':SOUR:SWE:SPAC LIN')                  # Select linear staircase sweep
        self.keithley.write(f':TRIG:COUN {self.nbr_mesures}')           # Trigger count = # sweep points
        self.keithley.write(":SOUR:DEL 0.1")
        #For single sweep, trigger count should equal number of points in sweep: Points =
        #(Stop-Start)/Step + 1. You can use SOUR:SWE:POIN? query to read the number
        #of points
        self.keithley.write(':OUTP ON')   # Turn on source output
        self.keithley.write(':READ?')   # Trigger sweep, request data
        self.mesures = self.keithley.read()
        self.keithley.write(':OUTP OFF')

        self.mesures = self.mesures.split(',')

        self.currents = []
        self.voltages = []

        for index in range(len(self.mesures)):
            if index %5 == 0:
                self.voltages.append(float(self.mesures[index]))
            elif index %5 == 1:
                self.currents.append(float(self.mesures[index]))

        self.currents = np.array(self.currents)
        self.voltages = np.array(self.voltages)

        # Perform linear regression
        self.slope, self.intercept, _, _, _ = linregress(self.currents, self.voltages)
        self.regression_line = self.slope * self.currents + self.intercept

        print("résistance mesurée: ", self.slope*4.532, "ohms")
        print(f"résistivité : {self.slope*4.532*self.sample_thickness*1e-4} ohm/cm")

        plt.plot(self.currents, self.voltages, marker='o', label="Measured Data")
        plt.plot(self.currents, self.regression_line, label=f"Fit: V = {self.slope:.3e}I + {self.intercept:.3e}", linestyle='--')
        plt.xlabel("Current (A)")
        plt.ylabel("Voltage (V)")
        plt.title("Current-Voltage Characteristic")
        plt.legend()
        plt.grid(True)
        plt.show()

        self.previous_state = self.state
        self.state = "SAVE"

    def save(self):
        #print("SAVE")

        if input("Enregistre la mesure ? (o/n) : ") == "o":
            self.tempFileNbr += 1
            self.csvFiles.append(f'{self.tempFileNbr}.csv')
            with open(f'{self.tempFileNbr}.csv', 'w') as csvfile:
                writer = csv.writer(csvfile, lineterminator='\n')
                writer.writerow(["","",""])
                writer.writerow([f"mesure # {self.tempFileNbr}", f"{self.sample_name}", ""])
                writer.writerow(["Resistance :", self.slope*4.532, "ohms"])
                writer.writerow(["Resistivite :", f"{self.slope*4.532*self.sample_thickness*1e-4}", "ohms/cm"])
                writer.writerow(["I (A)", "U (V)", ""])
                writer.writerows(
                    zip(self.currents, self.voltages)
                )

        if input("Recommencer la mesure ? (o/n) : ") == "o":
            if input("Garder les mêmes paramètres ? (o/n) : ") == "o":
                self.previous_state = self.state
                self.state = "MEASURE"
            else:
                self.previous_state = self.state
                self.state = "DEFINITION PARAMETRES"

        elif self.tempFileNbr != 0:
            all_dataframes = []
            for file in self.csvFiles:
                df = pd.read_csv(file,)  # Skip the first row (header info specific to each file)
                all_dataframes.append(df)

            concatenated = pd.concat(all_dataframes,axis = 1, ignore_index=True)
            concatenated.to_csv(f"{self.file_name}.csv", index=False)

            self.previous_state = self.state
            self.state = "STOP"

    def error(self):
        print("\n\n---ERREUR---\n\n")
        for message in self.error_messages:
            print(message)
        input("")
        self.state = "STOP"

    def stop(self):
        #print("STOP")

        for file in self.csvFiles:
            remove(file)

        self.keithley.close()

        exit()



if __name__ == "__main__":
    sm = StateMachine()
    sm.run()

