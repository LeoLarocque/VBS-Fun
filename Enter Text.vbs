Dim Message, input
input = InputBox("Enter Text: ", "Speak")
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak input
'The below code changes the voice to a female.'
'Dim msg, sapi'
'Set sapi = createObject("sapi.spvoice")'
'Set sapi.Voice = sapi.GetVoices.Item(1)'
'sapi.Speak "Hello world"'
'Lol = msgbox("Request done.")'

