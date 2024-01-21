from math import log10,sqrt

class antenna:
    def __init__(self, freq_min, freq_max, gain, distance):
        self.freq_min = freq_min
        self.freq_max = freq_max
        self.gain = int(gain*2)
        self.gs_min = 0
        self.gs_max = 0
        self.ac = 2
        self.at = 32.5 + (20 * log10(distance)) + (20 * log10(freq_min))

    def margem(self, min, max, gs_antenna):
        self.gs_max = self.ac + self.at + max - self.gain
        self.gs_min = self.ac + self.at + min - self.gain 
        if gs_antenna < self.gs_max and gs_antenna > self.gs_min:
            return float(gs_antenna + self.gain - self.ac - self.at)
        else:
            return 0
        
    def fresnel(self, distance):
        return str(sqrt(((3*(10**8)/(self.freq_max*(10**6)))*(distance/2)*(distance/2))/(distance/2 + distance/2))) 
