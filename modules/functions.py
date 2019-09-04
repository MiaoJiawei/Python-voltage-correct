import matplotlib.pyplot as plt
import numpy as np
import math

from scipy import interpolate


def maxvalue(voltage, current):
    res = []
    cur_max = max(current)
    vol_max = voltage[current.index(max(current))]
    res.append(vol_max)
    res.append(cur_max)
    return res


def leastsq(x_array, y_array):
    n = len(x_array)
    x_ave = sum(x_array) / n
    y_ave = sum(y_array) / n
    b_up_sum = 0
    b_sub_sum = 0
    for i in range(n):
        b_up_sum += (x_array[i])*(y_array[i])
        b_sub_sum += ((x_array[i])*(x_array[i]))

    b = (b_up_sum - n*x_ave*y_ave)/(b_sub_sum - n*x_ave*x_ave)
    a = y_ave - b*x_ave

    return b, a


def set_common_header(ws_raw, ws_res):
    ws_res.freeze_panes = 'A4'
    i = 1
    for raw in ws_raw.iter_rows(min_row=1, max_row=3, values_only=True):
        j = 1
        for cont in raw:
            ws_res.cell(i, j).value = cont
            j = j + 1
        i = i + 1
    return ws_res


def set_cutted_header(ws_raw, ws_cut):
    ws_cut.freeze_panes = 'A4'
    i = 1
    for raw in ws_raw.iter_rows(min_row=1, max_row=3, values_only=True):
        if i == 1 or i == 3:
            j = 1
            for cont in raw:
                if j % 2 == 1:
                    ws_cut.cell(i, 2*j-1).value = cont
                    ws_cut.cell(i, 2*j+1).value = cont
                else:
                    ws_cut.cell(i, 2*j-2).value = cont
                    ws_cut.cell(i, 2*j).value = cont
                j = j + 1
            i = i + 1
        elif i == 2:
            j = 1
            posi = 'posi'
            nega = 'nega'
            for cont in raw:
                if j % 2 == 1:
                    ws_cut.cell(i, 2*j).value = posi
                    ws_cut.cell(i, 2*j-1).value = posi
                else:
                    ws_cut.cell(i, 2*j).value = nega
                    ws_cut.cell(i, 2*j-1).value = nega
                j = j + 1
            i = i + 1
        else:
            break
    return ws_cut


def pushpack(ws, cont, col_index):
    j = 4
    for cell in cont:
        ws.cell(j, col_index).value = cell
        j = j + 1
    return ws


def resistor(worksheet):

    vol = []
    cur = []
    current_peaks = []
    voltage_peaks = []
    i = 0
    for col in worksheet.iter_cols(values_only=True):
        if i % 2 == 0:
            vol_temp = []
            j = 1
            for cell in col:
                if (j > 3) and (cell != None):
                    vol_temp.append(cell)
                j = j + 1
            vol = vol_temp
        else:
            cur_temp = []
            j = 1
            for cell in col:
                if (j > 3) and (cell != None):
                    cur_temp.append(cell)
                j = j + 1
            cur = cur_temp
            maxres = maxvalue(vol, cur)
            voltage_peaks.append(maxres[0])
            current_peaks.append(maxres[1])
        i = i + 1

    # Fit function.
    fit = leastsq(voltage_peaks, current_peaks)
    return [fit, [voltage_peaks, current_peaks]]


def cut(ws_raw, ws_cut):
    def cut(vol, cur):
        vol_posi = []
        cur_posi = []
        vol_nega = []
        cur_nega = []

        for i in range(len(vol)):
            if vol[i-1] < vol[i]:
                vol_posi.append(vol[i])
                cur_posi.append(cur[i])
            else:
                vol_nega.append(vol[i])
                cur_nega.append(cur[i])

        return[vol_posi, cur_posi, vol_nega, cur_nega]

    i = 0
    for col in ws_raw.iter_cols(values_only=True):
        if i % 2 == 0:
            vol_temp = []
            j = 1
            for cell in col:
                if (j > 3) and (cell != None):
                    vol_temp.append(cell)
                j = j + 1
            vol = vol_temp
        else:
            cur_temp = []
            j = 1
            for cell in col:
                if (j > 3) and (cell != None):
                    cur_temp.append(cell)
                j = j + 1
            cur = cur_temp
            cut_res = cut(vol, cur)
            vol_posi = cut_res[0]
            cur_posi = cut_res[1]
            vol_nega = cut_res[2]
            cur_nega = cut_res[3]
            pushpack(ws_cut, vol_posi, 2*i-1)
            pushpack(ws_cut, cur_posi, 2*i)
            pushpack(ws_cut, vol_nega, 2*i+1)
            pushpack(ws_cut, cur_nega, 2*i+2)
        i = i + 1
    return ws_cut


def correct(ws_raw, ws_cor, resistance):

    plt.figure(figsize=(8, 6))
    plt.grid(alpha=0.25)
    plt.xlabel('Voltage_correct')
    plt.ylabel('Current')
    max_vol = []
    min_vol = []
    i = 0
    for col in ws_raw.iter_cols(values_only=True):
        if i % 2 == 0:
            vol = []
            j = 1
            for cell in col:
                if (j > 3) and (cell != None):
                    vol.append(cell)
                j = j + 1
        else:
            cur = []
            j = 1
            for cell in col:
                if (j > 3) and (cell != None):
                    cur.append(cell)
                j = j + 1

            voltage_correct = []
            for k in range(len(vol)):
                voltage_correct.append(vol[k] - (cur[k] / resistance))
            plt.plot(voltage_correct, cur, label='corrected data')

            ws_cor = pushpack(ws_cor, voltage_correct, i)
            ws_cor = pushpack(ws_cor, cur, i+1)
            max_vol.append(max(voltage_correct))
            min_vol.append(min(voltage_correct))
        i = i + 1
    plt.legend(loc='upper left')
    plt.show()
    return ws_cor, max_vol, min_vol


def rearrangement(ws_raw, ws_res, voltage_range):
    def interpolation(array_raw, range):

        # Cut the range of voltage.
        i = 0
        index = []
        for vol in array_raw[0]:
            if (vol < range[0]) and (vol > range[1]):
                index.append(i)
            i = i+1

        index.reverse()

        for i in index:
            del array_raw[0][i]
            del array_raw[1][i]

        # Interpolation.
        vol = np.array(array_raw[0])
        cur = np.array(array_raw[1])
        inter_func = interpolate.interp1d(vol, cur)

        vol_inter = np.arange(range[0], range[1], 0.01)
        cur_inter = inter_func(vol_inter)

        array_res = [vol_inter, cur_inter]

        return array_res

    i = 0
    for col in ws_raw.iter_cols(values_only=True):
        if i % 2 == 0:
            vol = []
            j = 1
            for cell in col:
                if (j > 3) and (cell != None):
                    vol.append(cell)
                j = j + 1
        else:
            cur = []
            j = 1
            for cell in col:
                if (j > 3) and (cell != None):
                    cur.append(cell)
                j = j + 1

            array_raw = [vol, cur]
            inter = interpolation(array_raw, voltage_range)

            vol = inter[0]
            cur = inter[1]
            ws_res = pushpack(ws_res, vol, i)
            ws_res = pushpack(ws_res, cur, i+1)
        i = i + 1
    return ws_res


def fit_curve(ws_raw, ws_res):
    def write_row(ws, content, row_index):
        for j, cell in enumerate(content, start=1):
            ws.cell(row_index, j).value = cell
        return ws

    def fit_nondiff(scrt, cur):
        y = []
        x = []
        for i in range(len(scrt)):
            x.append(math.sqrt(scrt[i]))
            y.append((cur[i])/x[i])
        res = leastsq(x, y)
        k1 = res[0]
        res = []
        for v in scrt:
            res.append(k1*v)
        return res

    scrt = []
    for row in ws_raw.iter_rows(1, 1, values_only=True):
        for i, cell in enumerate(row):
            if i % 2 == 1:
                scrt.append(cell)

    for i, row in enumerate(ws_raw.iter_rows(values_only=True)):
        if i > 2:
            vol = []
            cur = []
            for j, cell in enumerate(row):
                if j % 2 == 0:
                    vol.append(cell)
                else:
                    cur.append(cell)

            scrt_p = []
            scrt_n = []
            cur_p = []
            cur_n = []
            for j, cur_v in enumerate(cur):
                if j % 2 == 0:
                    scrt_p.append(scrt[j])
                    cur_p.append(cur_v)
                else:
                    scrt_n.append(scrt[j])
                    cur_n.append(cur_v)
            cur_p = fit_nondiff(scrt_p, cur_p)
            cur_n = fit_nondiff(scrt_n, cur_n)

            cur_f = []
            for j in range(len(vol)):
                if j % 2 == 0:
                    cur_f.append(cur_p[int(j/2)])
                else:
                    cur_f.append(cur_n[int((j-1)/2)])

            content = []
            for j,vol_v in enumerate(vol):
                content.append(vol_v)
                content.append(cur_f[j])

            ws_res = write_row(ws_res, content, (i+1))

    return ws_res
