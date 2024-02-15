import speech_recognition as sr
import win32com.client
import webbrowser
from AppOpener import open
import spotipy
import json
import openai
import datetime
import pytz
import docx2pdf
import os
import io
import requests
from io import BytesIO
import PIL
from PIL import Image
openai.api_key="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"
speaker=win32com.client.Dispatch("SAPI.SpVoice")
timem=datetime.datetime.now(pytz.timezone('Asia/Kolkata'))
def saym(text):
    speaker.Speak(text)
def listenm():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold=1
        r.adjust_for_ambient_noise(source,duration=0.5)
        audio = r.listen(source)
        try:
            query=r.recognize_google(audio,language='en-in')
            return query
        except Exception as e:
            return
def openaim():
    saym("what can I do for you")
    text=listenm()
    # text=input()
    print(text)
    response = openai.Completion.create(
        model="gpt-3.5-turbo-instruct",
        prompt=text,
        max_tokens=4000,
        n=1
    )
    print(response['choices'][0]['text'])
    tm=listenm()
    print(tm)
    if "speak" in tm:
        saym(response['choices'][0]['text'])
def doctopdf():
    from docx2pdf import convert
    t = input("Please enter the address of the Word file(write default for default address): ")
    name=input("enter name of file: ")
    output=input("name of output file: ")
    if(t=="default"):
        source_file = fr"C:\Users\Jal Parikh\Documents\OneDrive\Documents\{name}.docx"
    else:
        source_file = fr"{t}\{name}.docx";
    output_file = fr"C:\Users\Jal Parikh\Downloads\{output}.pdf"
    convert(source_file, output_file)
    return
def openapps_and_webm(t):
    if "open" in t and "website" in t:
        splitm = str.split(t)
        webbrowser.open_new(f"https://www.{splitm[-1]}.com/")
        saym(f"opening .{splitm[-1]}")
        return
    if "data structures and algorithms" in t:
        webbrowser.open_new("https://courses.iiit.ac.in/course/view.php?id=4233")
        saym(f"opening dsa course in moodle")
    elif "information communication" in t:
        webbrowser.open_new("https://courses.iiit.ac.in/course/view.php?id=4205")
        saym(f"opening information communication course in moodle")
    elif "analog electronic circuits" in t:
        webbrowser.open_new("https://courses.iiit.ac.in/course/view.php?id=4299")
        saym(f"opening analog electronic circuits course in moodle")
    elif "linear algebra" in t:
        webbrowser.open_new("https://courses.iiit.ac.in/course/view.php?id=4234")
        saym(f"opening linear algebra course in moodle")
    elif "moodle" in t:
        webbrowser.open_new("courses.iiit.ac.in")
        saym("opening moodle")
    elif "mess" in t:
        webbrowser.open_new("mess.iiit.ac.in")
        saym("opening mess portal")
    elif "youtube" in t and "search" in t:
        saym("what do you want to search")
        print("what do you want to search: ")
        # search=input()
        search=listenm()
        webbrowser.open_new(f"https://www.youtube.com/results?search_query={search}")
        saym(f"opening {search}")
    elif "ims" in t or "i m s" in t or "attendance" in t:
        webbrowser.open_new("ims.iiit.ac.in")
        saym("opening attendance")
    elif "g p t" in t or "gpt" in t:
        webbrowser.open_new("https://chat.openai.com/")
        saym("opening chat G P T")
    elif "whatsapp" in t:
        saym("opening whatsapp")
        open("whatsapp",match_closest=True)
    elif "spotify" in t:
        saym("opening Spotify", )
        open("spotify",match_closest=True)
    elif "bing" in t:
        saym("opening bing")
        open("edge",match_closest=True)
    elif "google" in t:
        saym("opening google")
        open("google",match_closest=True)
    elif "this p c" in t or "file manager" in t:
        saym("opening this p c")
        open("file explorer",match_closest=True)
    elif "word" in t:
        saym("opening word")
        open("word",match_closest=True)
    elif "powerpoint" in t:
        saym("opening powerpoint")
        open("powerpoint",match_closest=True)
    elif "excel" in t:
        saym("opening excel")
        open("excel",match_closest=True)
    else:
        splitm = str.split(t)
        webbrowser.open_new(f"https://www.{splitm[-1]}.com/")
        saym(f"opening .{splitm[-1]}")
    return
def highbeats():
    try:
        username="31q6necdtg5zwnzfh36vqt3vghni"
        client_id="4c26ec3bf4f64cc58b55a65c87d693c6"
        client_secret="a664f192907d4b779b4eee5fd058db42"
        redirect_uri = 'http://google.com/callback/'
        oauth_object = spotipy.SpotifyOAuth(client_id, client_secret,redirect_uri)
        token_dict = oauth_object.get_access_token()
        token=token_dict['access_token']
        sp = spotipy.Spotify(auth=token)
        user_name=sp.current_user()
        saym("say the song you are looking for")
        song=listenm()
        #song=input()
        res=sp.search(song,1,0,"track")
        song_dict=res['tracks']
        song_items=song_dict['items']
        sm=song_items[0]['external_urls']['spotify']
        webbrowser.open(sm)
        return
    except Exception as e:
        return
def imagem():

    saym("what kind of image do you want to make")
    t=listenm()
    # t = input()
    namem=t.split()
    print(t)
    try:
        response = openai.Image.create(
            prompt=t,
            size="1024x1024",
            n=1,
        )

        um = response["data"][0]["url"]
        res = requests.get(um)
        imagem = res.content
        im = Image.open(BytesIO(imagem))
        im.save(fr"C:\Users\Jal Parikh\Downloads\{namem[0]}.png")
        print("image made")
        saym("image made")

    except Exception as e:
        print("could not make image")
    return

if __name__=='__main__':
    if(timem.hour<12 and timem.hour>4):
        saym("Good morning Jal")
    elif(timem.hour>=12 and timem.hour<17):
        saym("Good afternoon Jal")
    elif(timem.hour>=17 and timem.hour<20):
        saym("Good evening Jal")
    else:
        saym("Hello Jal")
    while True:
        t=listenm()
        #t=input()
        if(t!=None):
            t=t.lower()
            print(t)
            if "open" in t or "show" in t:
                openapps_and_webm(t)
            if "play" in t or "song" in t:
                highbeats()
            if(t=="exit"):
                saym("exiting program")
                break;
            if "jarvis" in t:
                openaim()
            if"doc to pdf" in t:
                doctopdf()
            if"play video" in t:
                moviem()
            if"make an image" in t:
                imagem()
