import logging
from pptx import Presentation
import os
import subprocess
from random import randint

from flask import Flask, render_template

from flask_ask import Ask, statement, question, session

from os import listdir
from os.path import isfile, join


app = Flask(__name__)

ask = Ask(app, "/")

logging.getLogger("flask_ask").setLevel(logging.DEBUG)

presentation_url="/home/pi/Projexa_Files/"

pptno=0

filenames=[]

slideno=0   
@ask.launch


def new_game():

    welcome_msg = render_template('welcome')
    onlyfiles = [f for f in listdir(presentation_url) if isfile(join(presentation_url, f))]
    filenames=onlyfiles
    return question(welcome_msg)


@ask.intent("OpenIntent", default={'ppt_index':0}, convert={'ppt_index':int})
def open_presentation(ppt_index):
    """
    numbers = [randint(0, 9) for _ in range(3)]

    round_msg = render_template('round', numbers=numbers)

    session.attributes['numbers'] = numbers[::-1]  # reverse

    return question(round_msg)
    """
    global pptno
    pptno=ppt_index
    global slideno
    slideno=0
    global filenames
    print filenames
    presentation_url_file=presentation_url+filenames[int(ppt_index)-1]
    s="libreoffice --show "+presentation_url_file+" --norestore --nolockcheck"
    print s
    a=subprocess.Popen(s,shell=True)
    return statement("Opened ")   


@ask.intent("CurrentDetailsIntent")
def current_details_intent():
    global slideno
    global pptno
    return statement("Current Slide Number is "+str(slideno))

@ask.intent("ListIntent")

def list_presentations():       
    onlyfiles = [f for f in listdir(presentation_url) if isfile(join(presentation_url, f))]
    print onlyfiles
    msg=""
    index=1
    global filenames
    filenames=onlyfiles
    for afile in onlyfiles:
        msg=msg+"To open "+afile+" . Say open "+str(index)+". "
        index=index+1
    return question(msg)

@ask.intent("CloseIntent")
def close_presentation():
    global slideno
    slideno=0
    print "something"
    a=subprocess.Popen( "xdotool search --name 'Impress' key Escape", shell=True)
    return statement("closed")

@ask.intent("NextIntent", default={'no_slides_right':1}, convert={'no_slides_right':int})
def next_slide(no_slides_right): 
    print "right"
    print str(no_slides_right)
    i=0
    
    global slideno
    s="xdotool search --name 'Impress' key --delay 1000 Right"
    for i in range(1,int(no_slides_right)):
        s=s+" Right"
        slideno=slideno+1
    a=subprocess.Popen( s, shell=True)
    return statement("Moved")

@ask.intent("PrevIntent", default={'no_slides_left':1}, convert={'no_slides_left':int})
def prev_slide(no_slides_left):
    print "left"
    print str(no_slides_left)
    i=0
    global slideno
    x=slideno-int(no_slides_left)
    if x>=0:
        s="xdotool search --name 'Impress' key --delay 1000 Left"
        for i in range(1,int(no_slides_left)):
            s=s+" Left"
            slideno=slideno-1
        a=subprocess.Popen( s, shell=True)
        return statement("Moved")
    else:
        return statement("Cannot Move sorry")
if __name__ == '__main__':

    app.run(debug=True)
