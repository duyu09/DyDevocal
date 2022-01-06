# Copyright(c) 2020-2021 Qilu University of Technology,Class of Software Engineering(Software Development) 21-1 Duyu
import sys
import numpy as np
from scipy.io import wavfile
from scipy import signal
import os
import pygame
from pydub import AudioSegment
import ffmpeg
def Play_Music(musicname,name):
    pygame.mixer.init()
    #os.system("cls")
    try:
       pygame.mixer.music.load(musicname)
    except:
       print("Error,Failed to load music.\n")
       
    os.system("title 正在播放：" + name)
    pygame.mixer.music.play()
    while(True):
       sys.stdout.write("正在播放："+name+"  -  "+ms_to_hours(pygame.mixer.music.get_pos())+'\r')
       sys.stdout.flush()
       if pygame.mixer.music.get_busy()==False:
           break

def ms_to_hours(millis):
    seconds, milliseconds = divmod(millis, 1000)
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    return ("%02d:%02d:%02d.%s" % (hours, minutes, seconds,str(milliseconds/100)[0:1]))

def MP3_To_WAV(mp3_path, wav_path):
    MP3_File = AudioSegment.from_mp3(file=mp3_path)
    MP3_File.export(wav_path, format="wav")

def SplitChannel(srcMusicFile):
   # read wav file
   sampleRate, musicData = wavfile.read(srcMusicFile)
   left = []
   right = []
   for item in musicData:
       left.append(item[0])
       right.append(item[1])
   mixed_data = np.array(left) - np.array(right)
   wavfile.write(srcMusicFile + "_output.wav", sampleRate, mixed_data)


os.system("cls")
os.system("chcp 437>nul")
print("Copyright(c)2021 齐鲁工业大学 计算机科学与技术学院 软件工程（软件开发）21-1班 DUYU\n")
print("DUYU多媒体音乐播放器\n")
while(True):
   os.system("title DUYU音乐播放器")
   try:
       a = sys.argv[1]
       ch = sys.argv[2]
   except:
       a = input("\n请输入音乐文件路径\n>")
       ch = input("请输入播放方式：0=正常播放，1=消音播放\n>")
   if a[0:1] == '"':
       a = a[1:len(a)-1]
   try:
       if ch == "1":
         print("请稍侯，正在处理\n")
         wavpa = os.getenv("temp") + r"\temp.wav"
         MP3_To_WAV(a,wavpa)
         SplitChannel(wavpa)
         os.remove(wavpa)
         Play_Music(wavpa + "_output.wav",a)
       else:
         Play_Music(a,a)
   except:
       print("The file '" + musicname + "' don't exist.")
       os.system("pause")
       sys.exit(0)
#End of DUYU Player Source File.
