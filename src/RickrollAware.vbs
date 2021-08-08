do
Dim message, sapi
message=("Never gonna give you up Never gonna let you down Never gonna run around and desert you Never gonna make you cry Never gonna say goodbye Never gonna tell a lie and hurt you")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message
loop