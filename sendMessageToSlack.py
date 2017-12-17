import os
from slackclient import SlackClient

BOT_ID = "inform_bot"
BOT_TOKEN = "xoxb-239480418918-VF6GMoNl0H4IdN1OpbVFSgF4"
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
