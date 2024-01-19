import speech_recognition as sr
import pyttsx3
import pyaudio
engine = pyttsx3.init()
alexa = sr.Recognizer()
try:
    with (sr.Microphone() as source2):
        alexa.adjust_for_ambient_noise(source2,duration=0.2);
        audio = alexa.listen(source2)
        text= alexa.recognize_google(audio)
        text=text.lower()
        print("You said: " + text)
 #       engine.say(text)
except sr.RequestError as e:
    print("could not interpret ,please speak again")
except sr.UnknownValueError:
    print("unknown speech")


