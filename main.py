#Programming with Python Project (Group Project - Voice Assistant)
import subprocess
import wolframalpha
import pyttsx3
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
import pytz
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from wikipedia.exceptions import DisambiguationError, PageError
from configparser import ConfigParser 
from colorama import init, Back, Fore 
 
init(autoreset=False)

print(Back.BLUE + Fore.WHITE)
  

 # Timeout after 5 seconds

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def wishMe():
	global assname
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("A Very Good Morning !")

	elif hour>= 12 and hour<18:
		speak("A Nice Winter Afternoon !") 

	else:
		speak("A pleasant Evening !") 

	assname =("Alexa 1 point o")
	speak("I am your Assistant")
	speak(assname)


def username():
	speak("What should i call you ?")
	uname = takeCommand()
	speak(" Warm Welcome Miss")
	speak(uname)
	columns = shutil.get_terminal_size().columns
	
	print("#####################".center(columns))
	print("Welcome Ms.", uname.center(columns))
	print("#####################".center(columns))
	
	speak("How can i Help you, mam")

def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		
		print("Listening...")
		r.pause_threshold = 1
		audio = r.listen(source)

	try:
		print("Recognizing...") 
		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")

	except Exception as e:
		print(e) 
		print("Unable to Recognize your voice.") 
		return "None"
	
	return query

def getRecipientEmail():
    speak("To whom should I send the email?")
    recipient = takeCommand()
    return recipient


def sendEmail():
    try:
        speak("What do you want to say?")
        content = takeCommand()
        speak("To whom should I send the email?, Please enter receiver's email id , below")
        to = input("Enter the recipient's email address: ")  # Replace this with a call to takeCommand for full voice input
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.login('radhikaarora366@gmail.com', 'vqjh oyiu rkkd wswm')
        server.sendmail('radhikaarora366@gmail.com', to, content)
        server.close()
        speak("Email has been sent!,Your work got much simplified")
    except Exception as e:
        print(e)
        speak("I am unable to send this email at the moment.Please check the Internet Connectivity !")


def play_music():
    music_dir = "C:\\Users\\radhi\\Music"
    songs = os.listdir(music_dir)

    # Filter to include only .mp3 files or other audio file types
    songs = [song for song in songs if song.endswith('.mp3')]
    if songs:
        # Randomly select a song
        song_to_play = random.choice(songs)
        speak(f"Oh Nice , Playing {song_to_play}")
        os.startfile(os.path.join(music_dir, song_to_play))
    else:
        speak("No music files found.")

# extract key from the 
# configuration file 


def NewsFromBBC():
     
    # BBC news api
    # following query parameters are used
    # source, sortBy and apiKey
    query_params = {
      "source": "bbc-news",
      "sortBy": "top",
      "apiKey": "929a777c57844204b8e9b249451050a3"
    }
    main_url = " https://newsapi.org/v1/articles"
 
    # fetching data in json format
    res = requests.get(main_url, params=query_params)
    open_bbc_page = res.json()
 
    # getting all articles in a string article
    article = open_bbc_page["articles"]
 
    # empty list which will 
    # contain all trending news
    results = []
     
    for ar in article:
        results.append(ar["title"])
         
    for i in range(len(results)):
         
        # printing all trending news
        print(i + 1, results[i])
 
    #to read the news out loud for us
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(results)  

def calculate(query):
    if not query or not isinstance(query, list) or all(not s.strip() for s in query):
        speak("Please provide what you want me to calculate.")
        query = takeCommand().split()  # Get additional input from the user
        if not query or all(not s.strip() for s in query):
            print("Error: Invalid or empty query provided.")
            return
    app_id = "T6RQV7-RUKPVP5ATU"
    client = wolframalpha.Client(app_id)
    try:
        res = client.query(' '.join(query))
        answer = next(res.results).text
        print("Answer:", answer)
        speak(f"The answer is {answer}")
    except Exception as e:
        print("An error occurred:", e)
        speak("Sorry, I couldn't process the calculation.")

if __name__ == '__main__':
	clear = lambda: os.system('cls')
	
	# This Function will clean any
	# command before execution of this python file
	clear()
	wishMe()
	username()
	
	while True:
		
		query = takeCommand().lower()
		
		# All the commands said by user will be 
		# stored here in 'query' and will be
		# converted to lower case for easily 
		# recognition of command
		if 'wikipedia' in query:
			speak('Searching Wikipedia....')
			webbrowser.open("https://www.wikipedia.org/")

		elif 'open youtube' in query:
			speak("Here you go to Youtube Enjoy \n")
			webbrowser.open("youtube.com")

		elif 'open google' in query:
			speak("Here you go to Google, Search bar\n")
			webbrowser.open("google.com")

		elif 'open stackoverflow' in query or "open stack overflow" in query:
			speak("Here you go to Stack Over flow, Happy coding")
			webbrowser.open("stackoverflow.com") 

		elif 'play music' in query or "play song" in query:
			speak("Seems you are in Nice mood today, Vibe with latest music")
			play_music()

		elif 'the time' in query or "time" in query or "current time" in query:
			ist = pytz.timezone("Asia/Kolkata")
			strTime = datetime.datetime.now(ist).strftime("%I:%M:%p") 
			speak(f"Mam, the current time in India is {strTime}")
			print({strTime})

		elif "date" in query or "current date" in query:
			strDate = datetime.datetime.now().strftime("%a %b %d %Z %Y")
			speak(f"Mam, the Date is {strDate}")
			print({strDate})

		elif 'open cisco' in query:
			codePath = r"C:\\Users\\radhi\\Desktop\\Cisco Packet Tracer.lnk"
			os.startfile(codePath)
			speak("Please check the task bar for your required App")

		elif 'email' in query:
			sendEmail()

		elif 'how are you' in query or 'how r u' in query :
			speak("I am fine, Thank you")
			speak("How are you, Mam")

		elif 'fine' in query or "good" in query or "i am also fine" in query:
			speak("It's good to know that your fine, Take Care weather is changing !")

		elif "change name" in query or "i want to rename you" in query or "i want to change your name" in query or "I want to rename " in query:
			speak("What would you like to call me, Sir ")
			assname = takeCommand()
			speak("Thanks for renaming me ! Such a nice name ")

		elif "what's your name" in query or "what is your name" in query or "whats your name" in query:
			speak("My friends call me")
			speak(assname)
			print("My friends call me", assname)

		elif 'exit' in query or "stop" in query or "stop working" in query:
			speak("Thanks for giving me your time,hope I added some value to you")
			exit()

		elif "who made you" in query or "who created you" in query: 
			speak("I have been created by Radhika and Pearl.")
			
		elif 'joke' in query or "tell me a joke" in query:
			speak(pyjokes.get_joke())
			speak ("Hope you Enjoyed !")
			
		elif "calculate" in query:
			speak("What would you like me to calculate?")
			speak("Dont give me a question on Calculas, I warn you !")
			detailed_query = takeCommand()
			if detailed_query:
				calculate(detailed_query.split())
			else:
				speak("I couldn't understand what you wanted to calculate.")


		elif "who i am" in query:
			speak("If you talk then definitely you are human.")

		elif "why you came to world" in query:
			speak("Thanks to Radhika and Pearl. further It's a secret")

		elif 'power point presentation' in query or 'ppt' in query:
			speak("opening Power Point presentation")
			power = r"C:\\Users\\radhi\\Downloads\\Sem 5\\DBMS\\Radhika 2230080.pptx"
			os.startfile(power)

		elif 'what is love' in query:
			speak("It is 7th sense that destroy all other senses")

		elif "who are you" in query:
			speak("I am your virtual assistant created by Radhika and Pearl")

		elif 'reason for you' in query:
			speak("I was created as a Minor project by Radhika and Pearl ")

		elif 'change background' in query:
			ctypes.windll.user32.SystemParametersInfoW(20, 0, r"C:\Users\radhi\Downloads\Projects\Python Project (Voice Assistant)\Wallpaper", 0)

			speak("Background changed successfully")

		elif "news" in query:
			NewsFromBBC() 
		
		elif 'lock window' in query or "lock my pc" in query:
			speak("locking the device")
			ctypes.windll.user32.LockWorkStation()

		elif 'shutdown system' in query or "shutdown" in query:
			speak("Hold On a Sec ! Your system is on its way to shut down")
			subprocess.run(['shutdown', '/s', '/f', '/t', '0'])

		elif 'empty recycle bin' in query:
			winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
			speak("Recycle Bin Recycled")

		elif "don't listen" in query or "stop listening" in query:
			speak("I will be on pause for next 5 minutes")
			time.sleep(5)
			

		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("You asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.nl/maps/place/" + location + " ")
			speak("Hoping , this was your desired location")

		elif "weather of" in query:
			query = query.replace("weather of", "")
			location = query
			speak("You asked to predict weather of ")
			speak(location)
			webbrowser.open("https://www.accuweather.com/en/in/ludhiana/205592/weather-forecast/205592")
			speak("Take care, some cold winds are expected !")

		elif "camera" in query or "take a photo" in query:
			try:
				speak("Initializing the camera, please hold still.")
				ec.capture(0, "Jarvis Camera", "img.jpg")
				speak("Photo captured successfully. Check your folder for the image.")
				speak("Expecting , it was a nice click !")
			except Exception as e:
				print(f"An error occurred: {e}")
				speak("I encountered an issue while trying to access the camera.")


		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "hibernate" in query or "sleep" in query or "hiber" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "log off" in query or "sign out" in query:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(5)
			subprocess.call(["shutdown", "/l"])

		elif "write a note" in query:
			speak("What should I write?")
			note = takeCommand()
			if note and note.strip() != "None":
				file = open("important_notes.txt", "a")
				speak("Should I include the date and time?")
				snfm = takeCommand().lower()
				if 'yes' in snfm or 'sure' in snfm:
					strTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
					file.write(f"{strTime} - {note}\n")
					speak("Note with date and time saved.")
				else:
					file.write(note + "\n")
					speak("Note saved without date and time.")
					file.close()
			else:
				speak("I couldn't understand what you said. Please try again.")

		elif "show note" in query or "read my note" in query:
			speak("Showing Notes")
			file = open("important_notes.txt", "r") 
			print(file.read())
			speak(file.read(6))

		elif "alexa" in query:
			wishMe()
			speak("Alexa 1 point o in your service Miss")
			speak(assname)

			
		elif "send message " in query:
				# You need to create an account on Twilio to use this service
				account_sid = 'Account Sid key'
				auth_token = 'Auth token'
				client = Client(account_sid, auth_token)

				message = client.messages \
								.create(
									body = takeCommand(),
									from_='Sender No',
									to ='Receiver No'
								)

				print(message.sid)

		elif "Good Morning" in query:
			speak("A warm " +query)
			speak("How are you Mister")
			speak(assname)

		# most asked question from google Assistant
		elif "will you be my gf" in query or "will you be my bf" in query: 
			speak("I'm not sure about, may be you should give me some time")

		elif "how are you" in query:
			speak("I'm fine, glad you me that")

		elif "i love you" in query:
			speak("It's hard to understand")

		