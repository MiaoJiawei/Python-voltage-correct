import matplotlib.pyplot as plt

def correct_func(voltage, current, resistance):
    voltage_correct = []
    for i in range(len(voltage)):
        voltage_correct.append(voltage[i] - (current[i] * resistance))
    return voltage_correct

voltage = [8.19,2.72,6.39,8.71,4.7,2.66,3.78]
current = [7.01,2.78,6.47,6.71,4.1,4.23,4.05]
resistance = 0.6134953432516345

'''
plt.figure(figsize=(8,6))
plt.grid(alpha=0.25)
plt.xlabel('Voltage')
plt.ylabel('Current')
plt.plot(voltage, current, label='raw data')
plt.plot(voltage_correct, current, label='corrected data')
plt.legend(loc='upper left')
plt.show()
'''