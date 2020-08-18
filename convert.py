# -*- coding: UTF-8 -*-
import win32com.client
import time
import os
import sys
import shutil

def ppt_to_mp4(ppt_path,mp4_target,duration,resolution = 720,frames = 24,quality = 60,timeout = 120):
    # status:Convert result. 0:failed. -1: timeout. 1:success.
    status = 0
    if ppt_path == '' or mp4_target == '':
        return status
    # start_tm:Start time
    start_tm = time.time()

    # Start converting
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    presentation = ppt.Presentations.Open(ppt_path,WithWindow=True)
    # CreateVideo() function usage: https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentation.createvideo
    presentation.CreateVideo(mp4_target,0,duration,resolution,frames,quality)
    while True:
        try:
            time.sleep(0.1)
            if time.time() - start_tm > timeout:
                # Converting time out. Killing the PowerPoint process(An exception will be threw out).
                os.system("taskkill /f /im POWERPNT.EXE")
                status = -1
                break
            if os.path.exists(mp4_target) and os.path.getsize(mp4_target) == 0:
                # The filesize is 0 bytes when convert do not complete.
                continue
            status = 1
            break
        except Exception as e:
            print('Error! Code: {c}, Message, {m}'.format(c = type(e).__name__, m = str(e)))
            break
    print("WORKGIN TIME",time.time()-start_tm)
    if status != -1:
        ppt.Quit()

    return status
    