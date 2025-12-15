"""
Legacy instrument drivers (oscilloscopes / AWGs / power supplies / switches / SG / triggers).
Kept separate so the main multimeter tool chain (dmm_driver + dmm_ui) stays lightweight.
"""

import math
import struct
import time
import re
from datetime import datetime

import numpy as np
import serial
import matplotlib.pyplot as plt
from scipy.io import savemat
import scipy.interpolate as intpl
from pyvisa import constants as pyconst

from dmm_driver import bATEinst_base


class instAWG(bATEinst_base):
    Equip_Type = "awg"
    AWG_MODE_DC = "DC"
    AWG_MODE_SIN = "SIN"

    def set_output(self, on=True):
        self.set_error("Function not implemented")

    def set_amplitude(self, v):
        self.set_error("Function not implemented")

    def set_offset(self, v):
        self.set_error("Function not implemented")

    def set_dc(self, v):
        self.set_mode("DC")
        self.set_offset(v)

    def set_mode(self, mode):
        self.set_error("Function not implemented")

    def set_freq(self, freq):
        self.set_error("Function not implemented")

    def set_impedance(self, z):
        self.set_error("Function not implemented")

    def set_offset_quick(self, v):
        self.set_offset(v)


class instOSC_DS1104(bATEinst_base):
    Model_Supported = ["DS1104"]

    def __init__(self):
        super().__init__("osc")
        self.VisaAddress = "USB0::0x05E6::0x2450::04542396::INSTR"

    def callback_after_open(self):
        self.set_visa_timeout_value(10000)

    def set_x(self, xscale=None, offset=None):
        self.x_write(
            ([":TIM:SCAL %6e" % xscale, "*OPC?"] if xscale else [])
            + ([":TIM:OFFS %6e" % offset, "*OPC?"] if offset is not None else [])
        )

    def set_y(self, ch, yscale=None, yoffset=None):
        self.x_write(
            ([":CHAN%d:SCAL %6e" % (ch, yscale), "*OPC?"] if yscale else [])
            + ([":CHAN%d:OFFS %6e" % (ch, yoffset), "*OPC?"] if yoffset is not None else [])
        )

    def measure(self):
        self.x_write(
            [
                ":STOP",
                "*OPC?",
                "SING",
                "$WAIT=1"
                ":TRIG:STAT?$LOOP UNTIL = STOP in 20",
            ]
        )

    def start(self):
        self.x_write([":RUN"])

    def load_setup(self, fn):
        self.x_write([":LOAD:SET '%s'" % fn, "*OPC?"])

    def save_image(self, fn):
        self.x_write([":STORage:IMAGe:TYPE PNG", "*OPC?"])
        self.write(":DISPlay:DATA?")
        dd = self.read_block()
        with open(fn, "wb") as fid:
            fid.write(dd)

    def save_waveform(self, fn):
        res = [
            k.strip() == "1"
            for k in self.x_write([":CHAN1:DISP?", ":CHAN2:DISP?", ":CHAN3:DISP?", ":CHAN4:DISP?"])
        ]
        chs = []
        for k in range(len(res)):
            if res[k]:
                chs.append(k + 1)
        vvs = []
        for ch in chs:
            self.x_write(
                [
                    ":WAV:SOUR CHAN%d" % ch,
                    ":WAV:MODE RAW",
                    ":WAV:FORM BYTE",
                ]
            )
            point, av, xinc, xor, xref, yinc, yor, yref = [
                float(k) for k in self.x_write(":WAV:PRE?")[0].split(",")[2:]
            ]
            dd = []
            for st in range(1, int(point) + 1, 125000):
                self.x_write(
                    [
                        ":WAV:STAR %d" % st,
                        ":WAV:STOP %d" % (st + 125000 - 1 if point >= st + 125000 - 1 else point),
                    ]
                )
                self.write(":WAV:DATA?")
                dd += self.read_block()
            vvs.append([(k - yor - yref) * yinc for k in dd])
        pp = min([len(k) for k in vvs])
        with open(fn, "wt") as fid:
            for tt in range(pp):
                fid.write(
                    "\t".join(
                        ["%g" % ((tt - xor - xref) * xinc)]
                        + ["%g" % vvs[k][tt] for k in range(len(chs))]
                    )
                    + "\n"
                )


class instOSC_MDO34(bATEinst_base):
    Model_Supported = ["MDO34"]

    def __init__(self):
        super().__init__("osc")
        self.VisaAddress = "USB::0x0699::0x052C::C050152::INSTR"

    def set_x(self, xscale=None, offset=None):
        self.x_write(
            ([":TIM:SCAL %6e" % xscale, "*OPC?"] if xscale else [])
            + ([":TIM:OFFS %6e" % offset, "*OPC?"] if offset is not None else [])
        )

    def set_y(self, ch, yscale=None, yoffset=None):
        self.x_write(
            ([":CHAN%d:SCAL %6e" % (ch, yscale), "*OPC?"] if yscale else [])
            + ([":CHAN%d:OFFS %6e" % (ch, yoffset), "*OPC?"] if yoffset is not None else [])
        )

    def measure(self):
        self.x_write([":STOP", "*OPC?", "SING", "$WAIT=1$"])

    def start(self):
        self.x_write([":RUN"])

    def load_setup(self, fn):
        self.x_write([":LOAD:SET '%s'" % fn, "*OPC?"])

    def save_image(self, fn):
        self.x_write([":SAV:IMAG:FILEF PNG", "SAV:IMAG " + fn, "*OPC?"])

    def save_waveform(self, fn):
        self.x_write(
            [
                ":WFMO:ENC BIN",
                ":WFMO:BN_FMT RI",
                ":WFMO:BYT_O MSB",
                ":WFMO:BYT_N 1",
                ":DAT:SOU CH1",
                ":DAT:START 1",
                ":DAT:STOP 20000000",
                "*OPC?",
            ]
        )
        pre = self.x_write(":WFMO?")[0].split(";")
        n = int(pre[6])
        x_inc = float(pre[10])
        x_off = float(pre[11])
        y_inc = float(pre[14])
        y_off = float(pre[15])
        n_block = 200000
        with open(fn, "wb") as fid:
            hd = struct.pack("5d", n, x_inc, x_off, y_inc, y_off)
            fid.write(hd)
            for k in range(1, n, n_block):
                self.x_write(
                    [
                        ":DAT:START %d " % k,
                        ":DAT:STOP %d" % (k + n_block - 1 if (k + n_block - 1) <= n else n),
                        "*OPC?",
                    ]
                )
                self.write("CURV?")
                dd = self.read_block()
                fid.write(dd)


class instOSC_DHO1204(bATEinst_base):
    Model_Supported = ["DHO1204"]

    def __init__(self):
        super().__init__("osc")
        self.VisaAddress = "USB::0x1AB1::0x0610::HDO1B254200719::INSTR"

    def callback_after_open(self):
        self.set_visa_timeout_value(8000)

    def set_x(self, xscale=None, offset=None):
        self.x_write(
            ([":TIM:SCAL %6e" % xscale, "*OPC?"] if xscale else [])
            + ([":TIM:OFFS %6e" % offset, "*OPC?"] if offset is not None else [])
        )

    def set_y(self, ch, yscale=None, yoffset=None):
        self.x_write(
            ([":CHAN%d:SCAL %6e" % (ch, yscale), "*OPC?"] if yscale else [])
            + ([":CHAN%d:OFFS %6e" % (ch, yoffset), "*OPC?"] if yoffset is not None else [])
        )

    def measure(self):
        self.x_write([":STOP", "*OPC?", "SING", "$WAIT=1$"])

    def start(self):
        self.x_write([":RUN"])

    def stop(self):
        self.x_write([":STOP"])

    def load_setup(self, fn):
        self.x_write([f":LOAD:SET '{fn}'", "*OPC?"])

    def save_image(self, fn):
        self.x_write([":SAV:IMAG:FILEF PNG", "*OPC?"])
        dd = self.read_block(":SAVE:IMAGe:DATA?")
        with open(fn, "wb") as fid:
            fid.write(dd)

    def set_acquire(self, depth=None, mode=None):
        if depth:
            self.x_write([f":ACQ:MDEP {depth:.0f}"])

    def save_waveform(self, fn, waves=None):
        if waves is None:
            chs = [
                (ch + 1) if self.query(f":CHAN{ch+1}:DISP?").startswith("1") else None
                for ch in range(4)
            ]
            chs = [k for k in chs if k]
            mm = {}
            for ch in chs:
                wv = self.read_waveform(ch)
                mm = {**mm, **wv}
            mm["channels"] = chs
        else:
            if isinstance(waves, dict):
                mm = waves
            elif isinstance(waves, list):
                mm = {}
                for wv in waves:
                    mm = {**mm, **wv}
                chs = [k["channels"][0] for k in waves]
                mm["channels"] = chs
        mm.pop("scale", None)
        mm.pop("data", None)
        savemat(fn, mm, appendmat=False)

    def read_waveform(self, ch):
        self.x_write(
            [
                ":STOP",
                "*OPC?",
                f":WAV:SOUR CHAN{ch}",
                ":WAV:MODE RAW",
                ":WAV:FORM BYTE",
                "*OPC?",
            ]
        )
        point = int(round(float(self.x_write([":ACQuire:MDEP?"])[0].strip())))
        self.x_write([":WAV:STAR 1", f":WAV:STOP {point}"])
        pre = self.x_write(":WAV:PRE?")[0]
        ff, tt, point, count, xinc, xor, xref, yinc, yor, yref = [float(k) for k in pre.split(",")]

        mm = {
            "xscale": [xinc, xor, xref],
            "xinc": xinc,
            "point": point,
            "format": ff,
            "count": count,
            "type": tt,
            "channels": [ch],
            "time": str(datetime.fromtimestamp(time.time())),
        }
        yscale = [yinc, yor, yref]

        mm[f"ch{ch}scale"] = yscale
        dd = []
        max_size = 125000
        for st in range(1, int(point) + 1, max_size):
            self.x_write(
                [
                    f":WAV:STAR {st}",
                    f":WAV:STOP {(st + max_size - 1 if point >= st + max_size - 1 else point)}",
                    "*OPC?",
                ]
            )
            self.write(":WAV:DATA?")
            dd += self.read_block()

        ddar = np.array(dd, dtype=np.uint8)
        mm[f"ch{ch}data"] = ddar
        mm["scale"] = yscale
        mm["data"] = ddar
        return mm

    def raw2float(self, raw, scale=None):
        if isinstance(raw, dict):
            scale = raw["scale"]
            raw = raw["data"]
        return (np.array(raw) - (scale[1] + scale[2])) * scale[0]

    def test(self):
        self.start()
        time.sleep(5)
        self.stop()
        plt.plot(self.raw2float(self.read_waveform(1)))
        plt.show()


class instSW_CP2102(bATEinst_base):
    Model_Supported = ["3000072"]

    def __init__(self):
        super().__init__("sw")
        self.VisaAddress = "COM7"
        self.get_cal_amp = None
        self.current_freq = 0

    def inst_open(self):
        rr = re.match(r"COM(\d+)", self.VisaAddress)
        if rr:
            self.VisaAddress = f"ASRL{rr[1]}::INSTR"
        return super().inst_open()

    def callback_after_open(self):
        self.Inst.set_visa_attribute(pyconst.VI_ATTR_ASRL_RTS_STATE, 0)

    def set_sw(self, on=None):
        self.check_open()
        if isinstance(on, str):
            on = on.lower() != "awg"
        self.Inst.set_visa_attribute(pyconst.VI_ATTR_ASRL_RTS_STATE, 0 if on else 1)
        time.sleep(0.2)

    def test(self):
        self.set_sw(1)
        time.sleep(1)
        self.set_sw(0)


class instSG_DSG836(bATEinst_base):
    Model_Supported = ["DSG836"]

    def __init__(self):
        super().__init__("sg")
        self.VisaAddress = "USB::0x1AB1::0x099C::DSG8M253400109::INSTR"
        self.get_cal_amp: intpl.interp1d = None
        self.current_freq = 0

    def calib_level(self, val):
        if self.get_cal_amp:
            return val / math.pow(10, float(self.get_cal_amp(self.current_freq)) / 20)
        else:
            return val

    def set_freq(self, freq):
        self.x_write([":FREQ %.2f" % freq, "*OPC?"])
        self.current_freq = freq

    def set_amp_v(self, amp_v):
        self.x_write([":LEV %.6fV" % self.calib_level(amp_v), "*OPC?"])

    def set_on(self, on=True):
        self.x_write([":OUTP %d" % (1 if on else 0), "*OPC?"])

    def set_lf_freq(self, freq):
        self.x_write([":LFO:FREQ %.2f" % freq, "*OPC?"])
        self.current_freq = freq

    def set_lf_amp_v(self, amp_v):
        self.x_write([":LFO:LEV %.6fV" % amp_v, "*OPC?"])

    def set_lf_shape(self, shape="SINE"):
        self.x_write([":LFO:SHAP SINE"])

    def set_lf_on(self, on=True):
        self.x_write([":LFO %d" % (1 if on else 0), "*OPC?"])


class instAWG_DG4102(bATEinst_base):
    Model_Supported = ["DG4102"]

    def __init__(self):
        super().__init__("awg")
        self.ch = 1
        self.VisaAddress = "USB::0x1AB1::0x0641::DG4E245103182::INSTRs"
        self.get_cal_level = None
        self.freqs = [0, 0]
        self.levels = None

    def callback_after_open(self):
        pass

    def calib_level(self, ch, val, freq=None):
        if self.get_cal_level:
            freq = self.freqs[ch - 1] if freq is None else freq
            return val / float(self.get_cal_level[ch - 1](freq))
        else:
            return val

    def sel_chan(self, ch):
        self.ch = ch

    class MODE:
        SIN = 1
        DC = 0
        PULSE = 2
        SQU = 3

    CH_ALL = 0

    def set_freq(self, freq, ch=None):
        if isinstance(freq, list):
            for ch, vv in enumerate(freq):
                self.x_write([":SOUR%d:FREQ %f" % (ch + 1, vv), "*OPC?"])
                self.freqs[ch] = vv
        else:
            for ch in self.ch2chs(ch):
                self.x_write([":SOUR%d:FREQ %f" % (ch, freq), "*OPC?"])
                self.freqs[ch - 1] = freq

    def ch2chs(self, ch):
        chs = self.ch if ch is None else ch
        if not isinstance(chs, list):
            chs = [chs]
        if chs == []:
            chs = [1, 2]
        return chs

    def set_reset(self):
        self.x_write(["*RST", "*OPC?", ":OUTP1:IMP INF", ":OUTP2:IMP INF"])

    def set_mode(self, mode=MODE.SIN, ch=None):
        if not isinstance(mode, str):
            mode = "PULSE" if mode == 2 else "DC" if mode == 0 else "SQU" if mode == 3 else "SIN"
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:APPL:%s" % (ch, mode), "*OPC?"])

    def set_sine_mode(self, freq=1e8, amp=0.01, ch=None):
        self.set_mode(self.MODE.SIN, ch)
        self.set_freq(freq, ch)
        self.set_amp(amp, ch)
        self.set_offset(0, ch)
        self.set_on(True, ch)

    def set_dc_mode(self, dc=0, ch=None):
        self.set_mode(self.MODE.SIN, ch)
        self.set_freq(1e-6, ch)
        self.set_amp(1e-3, ch)
        self.set_offset(dc, ch)
        self.set_on(True, ch)

    def set_phase(self, ph, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:PHAS:%s" % (ch, ph), "*OPC?"])

    def phase_sync(self, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:PHAS:INIT" % (ch), "*OPC?"])

    def set_amp(self, amp, ch=None):
        if isinstance(amp, list):
            for ch, vv in enumerate(amp):
                self.x_write([":SOUR%d:VOLT:AMPL %.4f" % (ch + 1, self.calib_level(ch + 1, vv)), "*OPC?"])
        else:
            for ch in self.ch2chs(ch):
                self.x_write([":SOUR%d:VOLT:AMPL %.4f" % (ch, self.calib_level(ch, amp)), "*OPC?"])

    def set_burst_phase(self, ph, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:BURS:PHAS %.4f" % (ch, ph), "*OPC?"])

    def set_offset(self, v, ch=None):
        if isinstance(v, list):
            for ch, vv in enumerate(v):
                self.x_write([":SOUR%d:VOLT:OFFS %.4f" % (ch + 1, self.calib_level(ch + 1, vv, 0)), "*OPC?"])
        else:
            for ch in self.ch2chs(ch):
                self.x_write([":SOUR%d:VOLT:OFFS %.4f" % (ch, self.calib_level(ch, v, 0)), "*OPC?"])

    def set_on(self, on=True, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":OUTP%d %d" % (ch, (1 if on else 0)), "*OPC?"])

    def set_data_rate_test(self, afreq=200e3, bfreq=100, bursts=500, level=3.3):
        self.x_write(
            [
                "*RST",
                "*OPC?",
                ":OUTP1:IMP INF",
                ":OUTP2:IMP INF",
                ":SOUR1:APPL:SQU %f, %.2f, %.2f,300" % (afreq, level / 2, level / 4),
                ":SOUR2:APPL:SQU %f, %.2f, %.2f,0" % (bfreq, level / 2, level / 4),
                "*OPC?",
                ":SOUR2:TRIG:SOUR MAN",
                ":SOUR2:BURS:MODE TRIG",
                "*OPC?",
                ":SOUR2:BURS:NCYC %d" % bursts,
                "*OPC?",
                ":SOUR2:BURS ON",
                ":OUTP1 1",
                ":OUTP2 1",
                "*OPC?",
            ]
        )

    def fire_burst_manul_trigger(self, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:BURS:TRIG" % ch, "*OPC?"])

    def reset(self):
        self.x_write(
            [
                ":SYST:PRES DEF",
                "*OPC?",
                ":OUTP1:IMP INF",
                ":OUTP2:IMP INF",
                "*OPC?",
                ":SOUR1:APPL:SIN",
                ":SOUR2:APPL:SIN",
                "*OPC?",
            ]
        )


class instAWG_DG852(instAWG_DG4102):
    Model_Supported = ["DG852"]

    def __init__(self):
        super().__init__()
        self.VisaAddress = "USB::0x1AB1::0x0646::DG8R262900659::INSTR"

    def set_reset(self):
        self.x_write(["*RST", "$WAIT=1500$", "*OPC?", ":OUTP1:LOAD 50", ":OUTP2:LOAD 50"])

    def set_amp(self, amp, ch=None):
        if isinstance(amp, list):
            for ch, vv in enumerate(amp):
                self.x_write([":SOUR%d:VOLT %.4f" % (ch + 1, self.calib_level(ch + 1, vv)), "*OPC?"])
        else:
            for ch in self.ch2chs(ch):
                self.x_write([":SOUR%d:VOLT %.4f" % (ch, self.calib_level(ch, amp)), "*OPC?"])

    def phase_sync(self, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:PHAS:SYNC" % (ch), "*OPC?"])

    def set_data_rate_test(self, afreq=200e3, bfreq=100, bursts=500, level=3.3):
        self.x_write(
            [
                "*RST",
                "$WAIT=1500$",
                "*OPC?",
                ":OUTP1:LOAD 50",
                ":OUTP2:LOAD 50",
                ":SOUR1:APPL:SQU %f, %.2f, %.2f,300" % (afreq, level / 2, level / 4),
                ":SOUR2:APPL:SQU %f, %.2f, %.2f,0" % (bfreq, level / 2, level / 4),
                "*OPC?",
            ]
        )
        self.x_write(
            [
                ":SOUR2:TRIG:SOUR MAN",
                ":SOUR2:BURS:MODE TRIG",
                "*OPC?",
                f":SOUR2:BURS:NCYC {bursts}",
                "*OPC?",
                ":SOUR2:BURS:STAT ON",
                ":OUTP1 1",
                ":OUTP2 1",
                "*OPC?",
            ]
        )

    def fire_burst_manul_trigger(self, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([f":TRIG{ch}", "*OPC?"])

    def reset(self):
        self.x_write(
            [
                "*RST",
                "$WAIT=1500$",
                "*OPC?",
                ":OUTP1:LOAD 50",
                ":OUTP2:LOAD 50",
                "*OPC?",
                ":SOUR1:APPL:SIN",
                ":SOUR2:APPL:SIN",
                "*OPC?",
            ]
        )


class instDC_KA3003P(bATEinst_base):
    Model_Supported = ["KA3003P"]

    def __init__(self):
        super().__init__("dc")
        self.VisaAddress = "ASRL7::INSTR"

    def measure_v(self):
        return float(self.x_write(["VOUT1?"])[0])

    def measure_i(self):
        try:
            return float(self.x_write(["IOUT1?"])[0])
        except Exception:
            return float(self.x_write(["IOUT1?"])[0])

    def measure_iv(self):
        return (self.measure_i(), self.measure_v())

    def set_v(self, v):
        self.x_write(["VSET1:%.2fV" % v])

    def set_i(self, v):
        self.x_write(["ISET1:%.3fV" % v])

    def set_on(self, on=True):
        self.x_write(["OUT%d" % (1 if on else 0)])
        time.sleep(0.5)

    def test(self):
        self.set_v(3.3)
        self.set_i(0.2)
        self.set_on(1)
        i = self.measure_i()
        v = self.measure_v()
        print((i, v))
        self.set_on(0)


class instTrigger(bATEinst_base):
    Model_Supported = [""]

    CMD_ABSIZE = 0
    CMD_CCSIZE = 1
    CMD_STATE = 2

    def __init__(self):
        super().__init__("pwm")
        self.VisaAddress = "COM4"

    def inst_open(self):
        if not self.Inst:
            if not self.VisaAddress:
                self.set_error("Equip Address has not been set!")
            try:
                self.Inst = serial.Serial(self.VisaAddress, 115200, timeout=1)
            except serial.SerialException:
                time.sleep(2)
                self.Inst = serial.Serial(self.VisaAddress, 115200, timeout=1)
        return self.Inst

    def trigger(self, afreq, bfreq, csize):
        acycles = round(1e8 / afreq) - 1
        bcycles = round(afreq / bfreq)
        absize = acycles + (bcycles << 16)
        self.send(self.CMD_CCSIZE, 0x0000_0000)
        self.send(self.CMD_ABSIZE, absize)
        self.send(self.CMD_CCSIZE, csize + 0x8000_0000)

    def wait_done(self, maxdelay=10):
        st = time.time()
        while time.time() - st < maxdelay:
            v = self.send(self.CMD_STATE, 0x0000_0000)
            if (v & 0x03 == 2):
                break
            time.sleep(0.1)
        return time.time() - st

    def send(self, cmd, value):
        self.check_open()
        for _ in range(2):
            if self.Inst.in_waiting > 0:
                self.Inst.read(self.Inst.in_waiting)
            self.Inst.write(list("SV".encode("utf-8")) + [0, cmd] + list(struct.pack("I", value)))
            ret = self.Inst.read(8)
            if len(ret) == 8 and ret[0] == "s":
                break
            self.Inst.write([0] * 256)
            time.sleep(0.5)
        if len(ret) != 8:
            self.set_error("return error")
        return struct.unpack("II", ret)[1]
