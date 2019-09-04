import numpy
from scipy.optimize import leastsq

# Request resistor value by least square.
def resistor(voltage, current):
    # Define function.
    def hypothesis_func(w, x):
        w1,w0 = w
        return w1*x + w0

    # Build error function.
    def error_func(w, x, y):
        return hypothesis_func(w, x) - y

    # Set initer value.
    voltage = numpy.array(voltage)
    current = numpy.array(current)
    w_init = [20, 1]
    # Fit function.
    fit = leastsq(error_func, w_init, args=(voltage, current))    
    w_fit = fit[0]
    return w_fit

# Train set.
voltage = [8.19,2.72,6.39,8.71,4.7,2.66,3.78]
current = [7.01,2.78,6.47,6.71,4.1,4.23,4.05]
resistance = resistor(voltage,current)[0]
print (resistance)