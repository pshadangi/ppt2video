#!/usr/bin/env python

#------------------------------------------------------------------------------
# Tool functionality:
#   1. Open the ppt and read the notes section in each slide
#   2. Create audio files from the notes of each slide
#   3. Note the duration of each audio file
#   4. Set the slide transition duration in each slide as per audio duration
#   5. Attach the audio of each slide to the respective slide
#   6. Export the slideset to a mp4 video
#------------------------------------------------------------------------------
# Tool Usage: 
#   1. Create a working directory. Place the PPT in the working directory.
#   2. Set the parameter configurations below
#   3. Tool creates <WORKDIR_Tutorial> folder within WORKING_DIRECTORY
#   4. Notes, audio files and video are created inside the <WORKDIR_Tutorial> folder
#------------------------------------------------------------------------------

#------------------------------------------------------------------------------
# PARAMETERS TO CONFIGURE
#------------------------------------------------------------------------------
WORKING_DIRECTORY = r"C:\Temp\WORKDIR"
PPT_FILE = "sample.pptx"
RESOLUTION = 480    # do not change. 480 is chosen as a balance between quality and size
#------------------------------------------------------------------------------

import os
import sys
import glob
import logging
import platform
import time
import win32com.client
import string
import subprocess
import wave
import contextlib
import logging
import traceback
try:
    from mutagen.mp3 import MP3
except:
    pass

logging.basicConfig(format='%(asctime)s %(levelname)-8s [%(filename)s:%(lineno)3d] %(message)s',
    datefmt='%Y-%m-%d:%H:%M:%S',
    level=logging.DEBUG)

def get_mp3_duration(mp3_file):
    audio = MP3(mp3_file)
    return audio.info.length+1

def get_wave_duration(wavfile):
    duration = 1
    with contextlib.closing(wave.open(wavfile,'r')) as f:
        frames = f.getnframes()
        rate = f.getframerate()
        duration = frames / float(rate)
        logging.info( '%s -> %f' % (wavfile, duration) )
    return duration+1

def get_audio_duration(audio_file):
    filename, extn = os.path.splitext(audio_file)
    if extn == '.wav':
        return get_wave_duration(audio_file)
    elif extn == '.mp3':
        return get_mp3_duration(audio_file)
    else:
        sys.exit('Unexpected audio_file extension - %s' % extn)

def generate_audio(notes_catalog_file):
    if (sys.version_info.major == 3) and (sys.version_info.minor > 5):
        logging.info ("generating audio from speech catalog")
        result = subprocess.run(['powershell.exe', '-ExecutionPolicy',  'Bypass', '-File', '.\\text2audio.ps1', notes_catalog_file], stdout=subprocess.PIPE)
        logging.info (result.stdout.decode('utf-8'))
    else:
        proc = subprocess.Popen(['powershell.exe', '-ExecutionPolicy',  'Bypass', '-File', '.\\text2audio.ps1', notes_catalog_file],stdout=subprocess.PIPE)
        stdout_value = proc.communicate()[0].decode('utf-8')
        print('stdout:', repr(stdout_value))
    
def read_ppt_notes(pptApp, filepath, tutorial_dir, notes_catalog_file):
    try:
        if not os.path.exists(tutorial_dir):
            os.mkdir(tutorial_dir)
        presentation = pptApp.Presentations.Open(filepath)
        nTotalSlides = presentation.Slides.Count
        logging.info ('No. of slides = %d' % nTotalSlides)

        vt = chr(13)
        notes_file_list = []
        for nSlide in range(1, 1+nTotalSlides):
            strNotes = presentation.Slides(nSlide).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            strNotes = strNotes.replace(vt, '\n')
            notes_file = tutorial_dir + '/notes_slide_%d.md' % nSlide
            f = open(notes_file, 'w')
            f.write(strNotes)
            f.close
            notes_file_list.append(notes_file)
        
        f = open(notes_catalog_file, 'w')
        f.write( ' '.join(notes_file_list) )
        f.close()
        
        presentation.Close()
        
    except Exception as e:
        print(e)
        traceback.print_exc()

def add_audio_set_timing_genvideo(pptApp, filepath, tutorial_dir, notes_catalog_file, audio_type = '.wav'):
    try:
        if not os.path.exists(tutorial_dir):
            os.mkdir(tutorial_dir)
        notes_files = []
        with open(notes_catalog_file, 'r') as f:
            txt = f.read()
            notes_files = txt.split()
        
        presentation = pptApp.Presentations.Open(filepath)
        nTotalSlides = presentation.Slides.Count
        logging.info ('No. of slides = %d' % nTotalSlides)
    
        for nSlide in range(1, 1+nTotalSlides):
            notes_file = notes_files[nSlide-1]
            audiofile = notes_file + audio_type
            x = get_audio_duration(audiofile)
            duration = round(x, 1) + 3  # round off to 1st decimal, add 3s buffer

            presentation.Slides(nSlide).SlideShowTransition.AdvanceOnClick = True
            presentation.Slides(nSlide).SlideShowTransition.AdvanceOnTime = True
            presentation.Slides(nSlide).SlideShowTransition.AdvanceTime = duration
            
            logging.info("audio file: %s, file check: %s" % (audiofile, os.path.exists(audiofile)))
            audiofile_fullpath = os.path.abspath(audiofile)
            oShp = presentation.Slides(nSlide).Shapes.AddMediaObject2(audiofile_fullpath, False, True, 5, 5, 5, 5)
            try:
                oEffect = presentation.Slides(nSlide).TimeLine.MainSequence.AddEffect(oShp, 0x53)
                oEffect.EffectInformation.PlaySettings.HideWhileNotPlaying = True
                oEffect.Timing.TriggerDelayTime = 1.5         # start animation with delay
            except Exception as e:
                print(e)

        filebase, extn = os.path.splitext(filepath)
        newfilename = filebase + '_mastered' + extn
        presentation.SaveAs(newfilename)
        
        # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
        # format: ppSaveAsMP4 = 39
        # presentation.SaveAs(filebase + '_mastered.mp4', 39)
        
        video_file = '%s_mastered_%s.mp4' % (filebase, RESOLUTION)
        presentation.CreateVideo( video_file, True, 1, RESOLUTION) #, _FramesPerSecond_, _Quality_ )
        
        time.sleep(2)   # give time for proper file export/closure
        
        if os.path.isfile(video_file):
            logging.info("%s - generation in progress" % video_file)
        else:
            logging.info("%s - unable to create" % video_file)
        
        logging.info("video generation is asynchronous. Please close the ppt manually once video is generated")
        logging.info("please close the ppt in powerpoint application manually after video is created")
        #presentation.Close()

    except Exception as e:
        print(e)
        traceback.print_exc()

def do_main():
    TOOLDIR = os.path.abspath (os.path.dirname(__file__))
    logging.info("TOOL DIR: %s" % TOOLDIR)

    os.chdir(TOOLDIR)
    
    filepath = os.path.join(WORKING_DIRECTORY, PPT_FILE)
    tutorial_dir = os.path.splitext(os.path.basename(filepath))[0] + r'_tutorial'
    notes_catalog_file = tutorial_dir + '/notes_files.md'
    
    logging.info("PPT DIR: %s" % WORKING_DIRECTORY)
    logging.info("Tutorial material at: %s" % tutorial_dir)
    logging.info("Notes for speech: %s" % notes_catalog_file)
   
    pptApp = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
    
    pptApp.Visible = True
    time.sleep(2)
    
    read_ppt_notes(pptApp, filepath, tutorial_dir, notes_catalog_file)
    generate_audio(notes_catalog_file)
    
    add_audio_set_timing_genvideo(pptApp, filepath, tutorial_dir, notes_catalog_file)
    
    time.sleep(1)   # pause before quiting else
    sys.exit(0)

if __name__ == "__main__":
    do_main()
