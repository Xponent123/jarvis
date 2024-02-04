import speech_recognition as sr
import win32com.client
import webbrowser
from AppOpener import open
import spotipy
import json

speaker=win32com.client.Dispatch("SAPI.SpVoice")
def saym(text):
    speaker.Speak(text)
def listenm():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold=1
        audio = r.listen(source)
        try:
            query=r.recognize_google(audio,language='en-in')
            return query
        except Exception as e:
            return "some error occured"
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
        saym(f"opening dsa course in moodle")
    elif "moodle" in t:
        webbrowser.open_new("courses.iiit.ac.in")
        saym("opening moodle")
    elif "mess" in t:
        webbrowser.open_new("mess.iiit.ac.in")
        saym("opening mess portal")
    elif "youtube" in t and "search" in t:
        print("what do you want to search: ")
        search=input();
        webbrowser.open_new(f"https://www.youtube.com/results?search_query={search}")
        saym(f"opening {search}")

    elif "attendance" in t:
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
    username="31q6necdtg5zwnzfh36vqt3vghni"
    client_id="4c26ec3bf4f64cc58b55a65c87d693c6"
    client_secret="a664f192907d4b779b4eee5fd058db42"
    redirect_uri = 'http://google.com/callback/'
    oauth_object = spotipy.SpotifyOAuth(client_id, client_secret,redirect_uri)
    token_dict = oauth_object.get_access_token()
    token=token_dict['access_token']
    sp = spotipy.Spotify(auth=token)
    user_name=sp.current_user()
    # print(json.dumps(user_name, sort_keys=True,indent=4))
    saym("say the song you are looking for")
    # song=listenm()
    song=input()
    res=sp.search(song,1,0,"track")
    song_dict=res['tracks']
    song_items=song_dict['items']
    sm=song_items[0]['external_urls']['spotify']
    webbrowser.open(sm)
    return

if __name__=='__main__':
    saym("Hello I am jarvis")
    while True:
        # t=listenm()
        t=input()
        t=t.lower()
        if "open" in t or "show" in t:
            openapps_and_webm(t)
        if "play" in t or "song" in t:
            highbeats()
        if(t.lower()=="exit"):
            saym("exiting program")
            break;

