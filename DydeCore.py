# DydeCore Python Source File.
import sys
import numpy as np
from scipy.io import wavfile
from scipy import signal

a=sys.argv[1]
print ("Processing:" + str(a) + "\n")
def SplitChannel(srcMusic,NAME):
   # read wav file
   SampleRate, Data = wavfile.read(srcMusic)
   print ("Sample Rate:" + str(SampleRate) + "\n")
   lc = []
   rc = []
   for point in Data:
       lc.append(point[0])
       rc.append(point[1])

   Mix = np.array(lc) - np.array(rc)
   wavfile.write(NAME + '_Output.wav', SampleRate, Mix)

SplitChannel(str(a),str(a))
print ("Finished.\n" )
