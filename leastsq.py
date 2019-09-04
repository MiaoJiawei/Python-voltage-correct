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
    return b,a

voltage = [-0.28265, -0.27002, -0.2442, -0.22024, -0.20224, -0.18022]
current = [5.90877, 11.58457, 26.42927, 46.75133, 81.11133, 172.37867]

print(leastsq(voltage, current))