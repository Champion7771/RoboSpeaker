import win32com.client as wincom

speak = wincom.Dispatch("SAPI.SpVoice")

while True:
    text = input("Type what you want me to say\n")
    if text == 'q':
        speak.speak("BYE BYE BYE")
        break       
    speak.speak(text)