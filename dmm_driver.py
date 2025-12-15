import time
import re

import pyvisa as visa
from pyvisa import constants as pyconst


class bATEinst_Exception(Exception):
    pass


class bATEinst_base(object):
    Equip_Type = "None"
    Model_Supported = ["None"]
    bATEinst_doEvent = None
    isRunning = False
    RequestStop = False
    VisaRM = None

    def __init__(self, name=""):
        self.Name = name
        self.VisaAddress = None
        self.Inst = None

    def __del__(self):
        self.close()

    def set_error(self, ss):
        raise bATEinst_Exception("Equip %s error:\n%s" % (self.Name, ss))

    @staticmethod
    def open_VisaRM():
        if not bATEinst_base.VisaRM:
            bATEinst_base.VisaRM = visa.ResourceManager()
        return bATEinst_base.VisaRM

    def isvalid(self):
        return True if self.VisaAddress else False

    def inst_open(self):
        if not self.Inst:
            if not self.VisaAddress:
                self.set_error("Equip Address has not been set!")
            try:
                self.Inst = bATEinst_base.open_VisaRM().open_resource(self.VisaAddress)
            except Exception:
                time.sleep(0.5)
                if not self.Inst:
                    self.Inst = bATEinst_base.open_VisaRM().open_resource(self.VisaAddress)
        return self.Inst

    def inst_close(self):
        if self.Inst:
            try:
                self.Inst.close()
            except Exception:
                pass
            self.Inst = None

    def callback_after_open(self):
        pass

    def set_visa_timeout_value(self, tmo):
        self.Inst.set_visa_attribute(pyconst.VI_ATTR_TMO_VALUE, tmo)

    def check_open(self):
        if not self.Inst:
            try:
                self.inst_open()
            except Exception as e:
                self.set_error("Can not open address:%s\ninfo:%s" % (str(self.VisaAddress), str(e)))
            self.callback_after_open()
        return self.Inst

    def close(self):
        try:
            if self.Inst:
                self.inst_close()
        except Exception:
            pass
        self.Inst = None

    def read(self):
        self.check_open()
        try:
            ss = self.Inst.read()
        except Exception as e:
            self.set_error("read error\n info:" + str(e))
        return ss

    def write(self, ss):
        self.check_open()
        try:
            if isinstance(ss, list):
                for k in ss:
                    self.Inst.write(k)
            else:
                self.Inst.write(ss)
        except Exception as e:
            self.set_error("Write error\n info:" + str(e))

    def query(self, ss):
        self.write(ss)
        return self.read()

    def delay(self, sec):
        time.sleep(sec)

    def x_write(self, vvs, chx=""):
        if isinstance(vvs, str):
            vvs = vvs.splitlines()
        res = []
        for cc in vvs:
            cc = cc.strip()
            if not cc:
                continue
            cc = cc.replace("$CHX$", chx)
            if re.match(r"\$WAIT *= *(\d+) *\$", cc):
                self.delay(int(re.match(r"\$WAIT *= *(\d+) *\$", cc).group(1)) / 1000)
            else:
                if "?" in cc:
                    res.append(self.query(cc))
                else:
                    self.write(cc)
        return res


class instMultimeter(bATEinst_base):
    Equip_Type = "mm"

    MM_MODE_V = "V"
    MM_MODE_I = "I"
    MM_MODE_R = "R"
    MM_MODE_R4 = "R4"
    MM_RANGE_AUTO = "AUTO"

    MM_AC = "AC"
    MM_DC = "DC"

    def __init__(self, name=""):
        super().__init__(name)
        self.current_mode = None
        self.current_ac_dc = None
        self.current_range = None

    def set_mode(self, mode=MM_MODE_V, ac_dc=MM_AC):
        if len(mode) <= 2:
            mode = "VOLT" if mode == "V" else "CURR" if mode == "I" else None
            if mode is None:
                raise ValueError("模式不符合要求")
        self.x_write([f"CONF:{mode}:{ac_dc}", "*OPC?"])
        self.current_mode = mode
        self.current_ac_dc = ac_dc

    def set_range(self, rng=MM_RANGE_AUTO):
        self.x_write([f"CONF:{self.current_mode}:{self.current_ac_dc} {rng}", "*OPC?"])
        self.current_range = rng

    def measure(self):
        return float(self.x_write(f"MEAS:{self.current_mode}:{self.current_ac_dc}? {self.current_range}")[0])

    def measure_quick(self):
        return self.measure()

    def measure_i(self):
        self.set_mode(self.MM_MODE_I)
        return self.measure()

    def measure_v(self):
        self.set_mode(self.MM_MODE_V)
        return self.measure()

    def measure_r(self):
        self.set_mode(self.MM_MODE_R)
        return self.measure()


class instKS_34461A(instMultimeter):
    def __init__(self, name="", visa_address=""):
        super().__init__(name)
        self.VisaAddress = visa_address
        self.sleep_time = None
        self.time_dur = None
        self.time_dur_unit = None
