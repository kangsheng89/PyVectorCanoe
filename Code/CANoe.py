"""API for setup/usage of Canoe COM Client interface.
"""
#--------------------------------------------------------------------------
# Standard library imports
import os
import sys
import subprocess
import win32com.client

import time



#Vector Canoe Class
class CANoe:

    def __init__(self):
        self.application = None
        #check if there is any instance of CANoe process
        output = subprocess.check_output('tasklist',shell=True)
        #if CANoe process is still available, kill the process
        if "CANoe32.exe" in str(output):
            os.system("taskkill /im CANoe32.exe /f 2>nul >nul")
            
        #re-dispatch object for CANoe Application
        self.application = win32com.client.Dispatch("CANoe.Application")


    def open_simulation(self,cfgname):
        #open CANoe simulation
        if (self.application != None):
            #check for valid file and it is *.cfg file
            if os.path.isfile(cfgname) and (os.path.splitext(cfgname)[1]==".cfg"):
                self.application.Open(cfgname)
            else:
                raise FileNotFoundError("Can't find CANoe cfg file")
        else:
            raise RuntimeError("CANoe Application is missing,unable to open simulation")

    def close_simulation(self):
        #close CANoe simulation
        if (self.application != None):
            self.stop_Measurement()
            self.application.Quit()
            
        #make sure the CANoe is close properly, otherwise enforce taskkill
        output = subprocess.check_output('tasklist',shell=True)

        if "CANoe32.exe" in str(output):
            os.system("taskkill /im CANoe32.exe /f 2>nul >nul")           

        self.application = None
            
    def start_Measurement(self):
        retry = 0
        retry_counter = 5
        #try to establish measurement within 20s timeout
        while not self.application.Measurement.Running and (retry < retry_counter):
            
            self.application.Measurement.Start()
            time.sleep(1)
            retry += 1
        if (retry == retry_counter):
            raise RuntimeWarning("CANoe start measuremet failed, Please Check Connection!")
        

    def stop_Measurement(self):
        if self.application.Measurement.Running:
            self.application.Measurement.Stop()
        else:
            pass
            
    def get_EnvVar(self,var):

        if (self.application != None):     
            result = self.application.Environment.GetVariable(var)
            return result.Value
        else:
            raise RuntimeError("CANoe is not open,unable to GetVariable")
        
    def set_EnvVar(self,var,value):
        result = None
        
        if (self.application != None):
            #set the environment varible
            result = self.application.Environment.GetVariable(var)
            result.Value = value
            
            checker = self.get_EnvVar(var)
            #check the environment varible is set properly?
            while (checker != value):
                checker = self.get_EnvVar(var)
            
                
        else:
            raise RuntimeError("CANoe is not open,unable to SetVariable")

    def get_SigVal(self, channel_num, msg_name, sig_name, bus_type = "CAN"):
        """
        @summary Get the value of a raw CAN signal on the CAN simulation bus
        @param channel_num - Integer value to indicate from which channel we will read the signal, usually start from 1,
                             Check with CANoe can channel setup.
        @param msg_name - String value that indicate the message name to which the signal belong. Check DBC setup.
        @param sig_name - String value of the signal to be read
        @param bus_type - String value of the bus type - e.g. "CAN", "LIN" and etc.
        @return The CAN signal value in floating point value.
                Even if the signal is of integer type, we will still return by
                floating point value.
        @exception None
        """
        if (self.application != None):     
            result = self.application.GetBus(bus_type).GetSignal(channel_num, msg_name, sig_name)
            return result.Value
        else:
            raise RuntimeError("CANoe is not open,unable to GetVariable")


