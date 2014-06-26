from comtypes.client import CreateObject
from comtypes.gen import SpeechLib

engine = CreateObject("SAPI.SpVoice")
engine.Rate = 1
stream = CreateObject("SAPI.SpFileStream")

import sys
inf = open(sys.argv[1])
for i,line in enumerate(inf):
	outfile = "output_%03d.wav"%i
	stream.Open(outfile, SpeechLib.SSFMCreateForWrite)
	engine.AudioOutputStream = stream
	sel=None
	engine.Volume = 100
	line = line.strip()
	if line == "--":
		s,t = "Sam","Silence"
		engine.Volume = 0
	else:
		s,t = line.split(":", 1)
	for v in engine.GetVoices():
		d=v.GetDescription()
		if s.lower() in d.lower():
			sel=v
			break
	if sel:
		engine.Voice = sel
	else:
		print("cant find voice")
		exit(1)
	engine.speak(t)
	stream.close()
