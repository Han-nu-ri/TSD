import os
from slackclient import SlackClient

BOT_ID = "inform_bot"
BOT_TOKEN = "xoxb-239480418918-S5qMohPJpkeafjbrD0dMwkJc"
CHANNEL = "#general"

# send message to the channel
def sendMessage(message):
	# get slack client
	slackClient =  SlackClient(BOT_TOKEN)
	# send slack message
	slackClient.api_call (
		"chat.postMessage",
		channel = CHANNEL,
		text = message
)
