#!/usr/bin/python3.6
"""
ICP data extraction script
Version 1.1

MIT License

Copyright (c) 2018 Thomas Jungers

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""
import os
import re
from collections import OrderedDict
import logging
import xlrd
import xlwt
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

__version__ = "1.1"

ELM_RE = re.compile(r'([A-Z][a-z]?) +([0-9.]+)')
MASSES = {'H': 1.00794, 'He': 4.002602, 'Li': 6.941, 'Be': 9.012182,
          'B': 10.811, 'C': 12.0107, 'N': 14.0067, 'O': 15.9994,
          'F': 18.9984032, 'Ne': 20.1797, 'Na': 22.98977, 'Mg': 24.305,
          'Al': 26.981538, 'Si': 28.0855, 'P': 30.973761, 'S': 32.065,
          'Cl': 35.453, 'Ar': 39.948, 'K': 39.0983, 'Ca': 40.078,
          'Sc': 44.95591, 'Ti': 47.867, 'V': 50.9415, 'Cr': 51.9961,
          'Mn': 54.938049, 'Fe': 55.845, 'Co': 58.9332, 'Ni': 58.6934,
          'Cu': 63.546, 'Zn': 65.409, 'Ga': 69.723, 'Ge': 72.64,
          'As': 74.9216, 'Se': 78.96, 'Br': 79.904, 'Kr': 83.798,
          'Rb': 85.4678, 'Sr': 87.62, 'Y': 88.90585, 'Zr': 91.224,
          'Nb': 92.90638, 'Mo': 95.94, 'Tc': 97.907216, 'Ru': 101.07,
          'Rh': 102.9055, 'Pd': 106.42, 'Ag': 107.8682, 'Cd': 112.411,
          'In': 114.818, 'Sn': 118.71, 'Sb': 121.76, 'Te': 127.6,
          'I': 126.90447, 'Xe': 131.293, 'Cs': 132.90545,
          'Ba': 137.327, 'La': 138.9055, 'Ce': 140.116,
          'Pr': 140.90765, 'Nd': 144.24, 'Pm': 144.912744,
          'Sm': 150.36, 'Eu': 151.964, 'Gd': 157.25, 'Tb': 158.92534,
          'Dy': 162.5, 'Ho': 164.93032, 'Er': 167.259, 'Tm': 168.93421,
          'Yb': 173.04, 'Lu': 174.967, 'Hf': 178.49, 'Ta': 180.9479,
          'W': 183.84, 'Re': 186.207, 'Os': 190.23, 'Ir': 192.217,
          'Pt': 195.078, 'Au': 196.96655, 'Hg': 200.59, 'Tl': 204.3833,
          'Pb': 207.2, 'Bi': 208.98038, 'Po': 208.982416,
          'At': 209.9871, 'Rn': 222.0176, 'Fr': 223.0197307,
          'Ra': 226.025403, 'Ac': 227.027747, 'Th': 232.0381,
          'Pa': 231.03588, 'U': 238.02891, 'Np': 237.048167,
          'Pu': 244.064198, 'Am': 243.061373, 'Cm': 247.070347,
          'Bk': 247.070299, 'Cf': 251.07958, 'Es': 252.08297,
          'Fm': 257.095099, 'Md': 258.098425, 'No': 259.10102,
          'Lr': 262.10969, 'Rf': 261.10875, 'Db': 262.11415,
          'Sg': 266.12193, 'Bh': 264.12473, 'Hs': 269.13411,
          'Mt': 268.13882}
ICON = """R0lGODlhEAAQAMZaAFWBuliDu1iDvFmEvFqFvFuFvVyGvV2
HvV2Hvl6HvV6Ivl+Ivl+Jv1+JwGCJvmGKv2KKwGKLv2KLwGOLv2OLwGOMwGSMwGSNwmiPwmuSxG+Vxn
CVxnOXxnWZyXmcyXyeynyfzIOkzYWlzoam0Iinz4up0Iup0Yyq0Y6s04+s1JSszJGt0pWtzZGu05mwz
pqwz5Wx1Zmx0puy0J2yz5600J600Zq115y22Jy32KK41aC52aK62qW616W826i82ai+3KnA3bDD2rDE
37LE27PF3LXF2bfH27rI2rbJ4rvK3b3L3rrM47zN5MDN3cPP3cjW6cnW6cvY6s3a69Xe6Nrk8N7n8uH
p8+Lp8+nv9vL1+v////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////
////////yH5BAEKAH8ALAAAAAAQABAAAAezgH+Cg4SFhjQpGhstM4aCQQ0VFicmFhUXSYU5CBYWHlpa
Hp0ERINHBJ0FVaBUBp0DToIgnRYhoKAftCV/SpwWEkxaNz1aSBSdAX8vE69ZVwMCWFauFhA1LMcWC1o
6nUBaAJ0RKi7MFhxaGZ0fWg+dEjJGCp0YTwALCwxQvgNTfx3alkSRQlBKlB8SLIgQVMSAgx1CIkqEQW
FAk0ExEqywwbEjiQM+CvFAgaOkyRFDHKkkFAgAOw=="""


class InvalidFileError(Exception):
    def __init__(self, message):
        self.message = message


class CalibrationError(Exception):
    def __init__(self, message):
        self.message = message


class Wavelength:
    def __init__(self, text):
        self.concentrations = []
        self.intensities = []
        self.calibrated = False
        match = ELM_RE.match(text)
        if not match:
            raise InvalidFileError(
                "Cannot read element wavelength {}".format(text))
        self.element = match.group(1)
        if self.element not in MASSES:
            raise InvalidFileError(
                "Cannot read element {}".format(self.element))
        self.wavelength = match.group(2)

    def add(self, concentration, intensity):
        if not self.calibrated:
            self.concentrations.append(concentration)
            self.intensities.append(intensity)
        else:
            raise CalibrationError(
                "Recalibration during analysis is not supported")

    def fit(self):
        if self.calibrated:
            raise CalibrationError("Calibration already performed")
        self.calibrated = True

        results, covar = np.polyfit(
            self.concentrations, np.mean(self.intensities, 1), 1, cov=True)
        self.slope = results[0]
        self.intercept = results[1]
        self.slope_sd = (covar[0][0] / 2)**.5
        self.intercept_sd = (covar[1][1] / 2)**.5
        self.r_squared = np.corrcoef(
            self.concentrations, np.mean(self.intensities, 1)
        )[0][1]**2

    def errors(self, relative=False, max=False):
        calc_conc = np.array([self.get_conc(i) for i in self.intensities])
        error = calc_conc - self.concentrations
        if relative:
            error = error / self.concentrations
        if max:
            return np.max(error[error != np.inf])
        else:
            return error

    def get_conc(self, intensities, details=False):
        if not self.calibrated:
            raise CalibrationError("Calibration not performed")

        int_avg = np.mean(intensities)
        int_sd = np.std(intensities, ddof=1)

        conc = (int_avg - self.intercept) / self.slope
        conc_sd = conc * np.sqrt(
            (int_sd**2 + self.intercept_sd**2) /
            (int_avg - self.intercept)**2 +
            self.slope_sd**2 / self.slope**2)

        if details:
            return int_avg, int_sd, conc, conc_sd
        else:
            return conc


class Sample:
    """A sample

    Attributes:
        data         Reference to the data object
        intensities  Intensities
        results      2D list: wavelength, [int_avg, int_sd, conc, conc_sd]
        means        2D list: element, [mean_conc, mean_conc_sd]
    """

    def __init__(self, data):
        self.data = data
        self.intensities = {}
        self.results = {}
        self.means = {}

    def add(self, wavelength, intensities):
        self.intensities[wavelength] = intensities

    def calc(self):
        for wl_label, wl in sorted(self.data.wavelengths.items()):
            self.results[wl_label] = wl.get_conc(
                self.intensities[wl_label], details=True)
        self.__mean()

    def has_wl(self, wavelength):
        return wavelength in self.intensities

    def ratio(self, elms):
        ratio = (self.means[elms[0]][0] / MASSES[elms[0]] /
                 (self.means[elms[1]][0] / MASSES[elms[1]]))
        ratio_sd = ratio * np.sqrt((self.means[elms[0]][1] /
                                    self.means[elms[0]][0])**2 +
                                   (self.means[elms[1]][1] /
                                    self.means[elms[1]][0])**2
                                   )
        return ratio, ratio_sd

    def __mean(self):
        """Calculate the mean of the concentration from all wavelengths.

        v1.1: the mean is now inverse-variance weighted.
        """
        for elm, elm_wls in self.data.elements.items():
            count = len(elm_wls)
            total_x_var = 0
            total_1_var = 0
            for wl_label in elm_wls:
                total_x_var += self.results[wl_label][2] \
                    / self.results[wl_label][3]**2
                total_1_var += 1 / self.results[wl_label][3]**2
            self.means[elm] = (total_x_var / total_1_var,
                               np.sqrt(count / total_1_var))


class ICPData:
    def __init__(self, filename):
        self.internal_std = None
        self.wavelengths = {}
        self.qc = []
        self.samples = OrderedDict()
        self.elements = {}
        self.ratios = []
        self.nreplicates = 0
        self.__open_wb(filename)
        self.__check_wb()
        self.__parse()
        self.__fit()

    def __fit(self):
        for wl_label, wl in self.wavelengths.items():
            if wl.element not in self.elements:
                self.elements[wl.element] = []
            self.elements[wl.element].append(wl_label)
            wl.fit()
        self.nreplicates = len(first(first(self.samples).intensities))

    def cell(self, title):
        if title == "intensities":
            return self.sheet.row_values(
                self.row_num, self.col_num[title])
        else:
            return self.sheet.cell_value(
                self.row_num, self.col_num[title])

    def __open_wb(self, filename):
        self.wb = xlrd.open_workbook(filename)

    def __check_wb(self):
        pass

    def __parse(self):
        self.__parse_titles()
        for self.row_num in range(self.row_num, self.sheet.nrows):
            if self.sheet.cell(self.row_num, 0).ctype != xlrd.XL_CELL_EMPTY:
                if self.cell('type') == "Blank":
                    self.__parse_blank()
                elif self.cell('wavelength') == self.internal_std:
                    self.row_num += 1
                    continue
                elif self.cell('type') == "Standard":
                    self.__parse_standard()
                elif self.cell('type') == "Sample":
                    if self.cell('label').startswith('QC'):
                        self.__parse_qc()
                    else:
                        self.__parse_sample()
                self.row_num += 1

    def __parse_titles(self):
        self.sheet = self.wb.sheet_by_index(0)
        for i in range(0, 15):
            row = self.sheet.row_values(i, 0)
            if "Solution Label" in row:
                self.row_num = i + 1
                self.col_num = {}
                for index, title in enumerate(row):
                    if title == "Solution Label":
                        self.col_num['label'] = index
                    elif title == "Type":
                        self.col_num['type'] = index
                    elif title == "Element":
                        self.col_num['wavelength'] = index
                    elif title == "Flags":
                        self.col_num['flags'] = index
                    elif title == "Conc":
                        self.col_num['conc'] = index
                    elif title == "Time":
                        self.col_num['time'] = index
                    elif title == "Replicates (intensity)":
                        self.col_num['intensities'] = index
                break
        if not self.col_num:
            raise InvalidFileError("Cannot find data start")

    def __parse_blank(self):
        if self.cell("conc") == 1:
            self.internal_std = self.cell("wavelength")
        else:
            wl = self.cell("wavelength")
            self.wavelengths[wl] = Wavelength(wl)
            self.wavelengths[wl].add(0, self.cell("intensities"))

    def __parse_standard(self):
        match = ELM_RE.match(self.cell("label"))
        if not match:
            raise InvalidFileError("Cannot read standard label")
        elm = match.group(1)
        conc = float(match.group(2))
        wl = self.wavelengths[self.cell("wavelength")]
        if wl.element == elm:
            wl.add(conc, self.cell("intensities"))

    def __parse_qc(self):
        match = ELM_RE.match(self.cell("label")[3:])
        if not match:
            raise InvalidFileError("Cannot read QC label")
        elm = match.group(1)
        conc = match.group(2)
        wl = self.wavelengths[self.cell("wavelength")]
        if wl.element == elm:
            self.qc.append((self.cell("wavelength"),
                            conc, self.cell("time")))

    def __parse_sample(self):
        label = self.cell("label")
        i = 1
        while (label in self.samples and
               self.samples[label].has_wl(self.cell("wavelength"))):
            i += 1
            label = "{label}_{i}".format(label=self.cell("label"), i=i)

        if label not in self.samples:
            self.samples[label] = Sample(self)
        self.samples[label].add(self.cell("wavelength"),
                                self.cell("intensities"))


class App:
    def __init__(self, root):
        self.root = root
        self.wl_frames = {}

        np.seterr(divide='ignore')
        logging.basicConfig(filename='ICP_data.log', level=logging.ERROR,
                            format='%(levelname)s:%(asctime)s %(message)s')
        matplotlib.use("TkAgg")

        self.menu = tk.Menu(self.root)
        self.menu.add_command(label="Open", command=self.openFile)
        self.menu.add_command(label="Set ratios", command=self.set_ratios)
        self.menu.add_command(label="Make report", command=self.make_report)
        self.menu.add_command(label="Exit", command=self.root.quit)
        self.menu.entryconfig(2, state=tk.DISABLED)
        self.menu.entryconfig(3, state=tk.DISABLED)
        self.root.config(menu=self.menu)

    def remove_wl(self, wl_label):
        self.wl_frames[wl_label].grid_remove()
        wl = self.data.wavelengths[wl_label]
        elm = self.data.elements[wl.element]
        elm.remove(wl_label)
        if not elm:
            del self.data.elements[wl.element]
        del self.data.wavelengths[wl_label]

    def set_ratios(self):
        RatioWindow(self.root, self.data)

    def make_report(self):
        in_path = os.path.dirname(self.filename)
        save_path = tk.filedialog.asksaveasfilename(
            defaultextension=".xls", initialdir=in_path,
            initialfile=os.path.basename(in_path) + "_report",
            title="Report file", filetypes=[("Excel file", "*.xls")])
        if save_path:
            report_wb = xlwt.Workbook()
            bold_style = xlwt.easyxf('font: bold true')
            num_fmt = xlwt.easyxf(num_format_str='0.000')
            int_fmt = xlwt.easyxf('font: color gray50',
                                  num_format_str='0.00E+0')
            r2_fmt = xlwt.easyxf(num_format_str='0.00000')

            # calibration
            sheet = report_wb.add_sheet('calibration')
            sheet.write(1, 0, "y-intercept", bold_style)
            sheet.write(2, 0, "slope", bold_style)
            sheet.write(3, 0, "SD y-intercept", bold_style)
            sheet.write(4, 0, "SD slope", bold_style)
            sheet.write(5, 0, "r²", bold_style)

            col_num = 0
            for wl_label, wl in self.data.wavelengths.items():
                col_num += 1
                sheet.write(0, col_num, wl_label, bold_style)
                sheet.write(1, col_num, wl.intercept, num_fmt)
                sheet.write(2, col_num, wl.slope, num_fmt)
                sheet.write(3, col_num, wl.intercept_sd, num_fmt)
                sheet.write(4, col_num, wl.slope_sd, num_fmt)
                sheet.write(5, col_num, wl.r_squared, r2_fmt)

            sheet.write(7, 0,
                        ("Each intensity is the average of {} replicates"
                         ).format(self.data.nreplicates))
            if self.data.internal_std:
                sheet.write(
                    8, 0,
                    ("Intensities corrected by use of an internal standard: {}"
                     ).format(self.data.internal_std))

            sheet.write(10, 0,
                        ("ICP_data.pyw version {}"
                         ).format(__version__))

            # details data
            sheet = report_wb.add_sheet('data')
            sheet.write(0, 0, "Sample", bold_style)
            col_num = 0
            for wl_label in sorted(self.data.wavelengths):
                sheet.write(0, col_num + 1, wl_label + " I", bold_style)
                sheet.write(0, col_num + 2, "SD I", bold_style)
                sheet.write(0, col_num + 3, "Conc", bold_style)
                sheet.write(0, col_num + 4, "SD Conc", bold_style)
                col_num += 4

            row_num = 0
            for sample_name, sample in self.data.samples.items():
                row_num += 1
                sheet.write(row_num, 0, sample_name)
                col_num = 0
                sample.calc()
                for wl_label, wl in sorted(self.data.wavelengths.items()):
                    results = sample.results[wl_label]
                    sheet.write(row_num, col_num + 1,
                                results[0], int_fmt)
                    sheet.write(row_num, col_num + 2,
                                results[1], int_fmt)
                    sheet.write(row_num, col_num + 3,
                                results[2], num_fmt)
                    sheet.write(row_num, col_num + 4,
                                results[3], num_fmt)
                    col_num += 4

            # results
            sheet = report_wb.add_sheet('results')
            sheet.write(0, 0, "Sample", bold_style)
            col_num = 0
            for elm in sorted(self.data.elements):
                sheet.write(0, col_num + 1, "C({})".format(elm), bold_style)
                sheet.write(0, col_num + 2, "SD C({})".format(elm), bold_style)
                col_num += 2

            for ratio in self.data.ratios:
                sheet.write(0, col_num + 1,
                            "ratio {}/{}".format(ratio[0], ratio[1]),
                            bold_style)
                sheet.write(0, col_num + 2,
                            "SD ratio {}/{}".format(ratio[0], ratio[1]),
                            bold_style)
                col_num += 2

            row_num = 0
            for sample_name, sample in self.data.samples.items():
                row_num += 1
                sheet.write(row_num, 0, sample_name)
                col_num = 0
                for elm in sorted(self.data.elements):
                    sheet.write(row_num, col_num + 1,
                                sample.means[elm][0], num_fmt)
                    sheet.write(row_num, col_num + 2,
                                sample.means[elm][1], num_fmt)
                    col_num += 2

                for elms in self.data.ratios:
                    ratio, ratio_sd = sample.ratio(elms)
                    sheet.write(row_num, col_num + 1, ratio, num_fmt)
                    sheet.write(row_num, col_num + 2, ratio_sd, num_fmt)
                    col_num += 2

            # save report
            report_wb.save(save_path)
            messagebox.showinfo("Success", "The report was saved")

    def openFile(self):
        self.filename = None
        self.filename = filedialog.askopenfilename(
            parent=self.root, title="Select data file",
            filetypes=[("ICP data file", "*.xls")])
        if self.filename:
            self.handleFile()

    def handleFile(self):
        self.data = ICPData(self.filename)
        elms = []
        row_i = -1
        col_i = -1
        for wl_label, wl in sorted(self.data.wavelengths.items()):
            if wl.element not in elms:
                row_i += 1
                elms.append(wl.element)
                col_i = -1
            col_i += 1

            self.wl_frames[wl_label] = WlFrame(self, wl_label)
            self.wl_frames[wl_label].grid(row=row_i, column=col_i,
                                          sticky=tk.NSEW)
        grid_size = self.root.grid_size()
        for i in range(0, grid_size[0]):
            self.root.grid_columnconfigure(i, weight=1)
        for i in range(0, grid_size[1]):
            self.root.grid_rowconfigure(i, weight=1)
        self.menu.entryconfig(2, state=tk.NORMAL)
        self.menu.entryconfig(3, state=tk.NORMAL)


class WlFrame(tk.Frame):
    def __init__(self, app, wl_label):
        tk.Frame.__init__(self, app.root)
        info_frame = tk.Frame(self)

        wl = app.data.wavelengths[wl_label]
        max_err = wl.errors(relative=True, max=True) * 100

        f = Figure()
        f.subplots_adjust(left=0, right=1, bottom=0, top=1)
        a = f.add_subplot(111)
        a.axis('off')
        a.plot(wl.concentrations, np.mean(wl.intensities, 1), '.')
        a.plot([0, wl.concentrations[-1]],
               [wl.intercept,
                wl.intercept + wl.slope *
                wl.concentrations[-1]],
               'g-' if max_err < 5 else 'r-')

        tk.Label(
            info_frame, text="{} - {} nm".format(wl.element,
                                                 wl.wavelength)
        ).grid(row=0, column=0, columnspan=3)
        tk.Label(
            info_frame, text="r² ="
        ).grid(row=1, column=0, sticky=tk.E)
        tk.Label(
            info_frame, text="{:.5f}".format(wl.r_squared)
        ).grid(row=1, column=1, sticky=tk.E)
        tk.Label(info_frame, text="max error:").grid(
            row=2, column=0, sticky=tk.E)
        tk.Label(
            info_frame, text="{:.2f} %".format(max_err)
        ).grid(row=2, column=1, sticky=tk.E)
        tk.Button(
            info_frame, text="Remove",
            command=lambda wl=wl_label: app.remove_wl(wl)
        ).grid(row=1, column=2, rowspan=2)
        info_frame.pack()

        c = FigureCanvasTkAgg(f, self)
        c.show()
        c.get_tk_widget().pack()


class RatioWindow(simpledialog.Dialog):
    def __init__(self, root, data):
        self.elements = data.elements
        self.ratios = data.ratios
        simpledialog.Dialog.__init__(self, root, title="Ratios")

    def body(self, master):
        self.r_list = tk.Listbox(master)
        self.r_list.pack(expand=True, fill=tk.BOTH)

        for ratio in self.ratios:
            self.r_list.insert(tk.END, "{}/{}".format(ratio[0], ratio[1]))

    def buttonbox(self):
        box = tk.Frame(self)

        tk.Button(box, text="OK", width=10,
                  command=self.ok).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(box, text="Add...", width=10,
                  command=self.add).pack(side=tk.LEFT, padx=5, pady=5)
        tk.Button(box, text="Delete", width=10,
                  command=self.delete).pack(side=tk.LEFT, padx=5, pady=5)

        box.pack()

    def add(self):
        AddRatioWindow(self)

    def delete(self):
        index = self.r_list.index(tk.ACTIVE)
        del self.ratios[index]
        self.r_list.delete(index)


class AddRatioWindow(simpledialog.Dialog):
    def __init__(self, parent):
        simpledialog.Dialog.__init__(self, parent, title="Add a ratio")

    def body(self, master):
        tk.Label(master, text="First element").grid(row=0, column=0)
        tk.Label(master, text="Second element").grid(row=1, column=0)

        self.elm1 = tk.StringVar()
        self.elm1.trace('w', lambda name, index, mode,
                        sv=self.elm1: self.elm_fmt(sv))
        self.elm1_ent = tk.Entry(master, textvariable=self.elm1)
        self.elm1_ent.grid(row=0, column=1)
        self.elm2 = tk.StringVar()
        self.elm2.trace('w', lambda name, index, mode,
                        sv=self.elm2: self.elm_fmt(sv))
        self.elm2_ent = tk.Entry(master, textvariable=self.elm2)
        self.elm2_ent.grid(row=1, column=1)

        return self.elm1_ent

    def apply(self):
        elm1 = self.elm1.get()
        elm2 = self.elm2.get()

        self.parent.ratios.append((elm1, elm2))
        self.parent.r_list.insert(tk.END, "{}/{}".format(elm1, elm2))

    def validate(self):
        elms = self.parent.elements
        elm1 = self.elm1.get()
        elm2 = self.elm2.get()

        if not elm1 or not elm2:
            messagebox.showwarning(
                "Bad input",
                "Please provide both elements")
            return False
        if elm1 not in elms:
            messagebox.showwarning(
                "Bad input",
                "Element {} was not analysed".format(elm1))
            return False
        if elm2 not in elms:
            messagebox.showwarning(
                "Bad input",
                "Element {} was not analysed".format(elm2))
            return False
        if ((elm1, elm2) in self.parent.ratios or
                (elm2, elm1) in self.parent.ratios):
            messagebox.showwarning(
                "Bad input",
                "This ratio was already added")
            return False
        return True

    def elm_fmt(widget, var):
        elm = var.get()
        try:
            elm = elm[0].upper() + elm[1:2].lower()
            var.set(elm)
        except IndexError:
            pass


class ErrorCatcher:
    def __init__(self, func, subst, widget):
        self.func = func
        self.subst = subst
        self.widget = widget

    def __call__(self, *args):
        try:
            if self.subst:
                args = self.subst(*args)
            return self.func(*args)
        except SystemExit as msg:
            raise SystemExit(msg)
        except Exception as e:
            log_except(e)


def first(s):
    return next(iter(s.values()))


def log_except(e):
    logging.exception('')
    messagebox.showerror(type(e).__name__, e)


if __name__ == "__main__":
    try:
        tk.CallWrapper = ErrorCatcher
        root = tk.Tk()
        root.title("ICP data extractor")
        icon = tk.PhotoImage(data=ICON)
        root.tk.call('wm', 'iconphoto', root._w, icon)
        app = App(root)
        root.mainloop()
    except Exception as e:
        log_except(e)
