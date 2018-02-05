import requests
import json
import time, datetime, openpyxl
from openpyxl.workbook import Workbook

from requests_oauthlib import OAuth1

# prepping the datetime for use later
dt = str(datetime.datetime.now().strftime ("%Y%m%d %H%M%S"))

# authentication pieces
client_key    = ""
client_secret = ""
token         = ""
token_secret  = ""

# the base for all Twitter calls
base_twitter_url = "https://api.twitter.com/1.1/"

# setup authentication
oauth = OAuth1(client_key,client_secret,token,token_secret)


#
# Download Tweets from a user profile
#
def download_tweets(screen_name,number_of_tweets,max_id=None):
    
    api_url  = "%s/statuses/user_timeline.json?" % base_twitter_url
    api_url += "screen_name=%s&" % screen_name
    api_url += "count=%d" % number_of_tweets
    
    if max_id is not None:
        api_url += "&max_id=%d" % max_id

    # send request to Twitter
    response = requests.get(api_url,auth=oauth)
    
    if response.status_code == 200:
        
        tweets = json.loads(response.content)
        
        return tweets
    

    return None

#
# Takes a username and begins downloading all Tweets
#
def download_all_tweets(username):
    full_tweet_list = []
    max_id          = 0
    
    # grab the first 200 Tweets
    tweet_list   = download_tweets(username,200)
    
    # grab the oldest Tweet
    oldest_tweet = tweet_list[-1]
    
    # continue retrieving Tweets
    while max_id != oldest_tweet['id']:
    
        full_tweet_list.extend(tweet_list)

        # set max_id to latest max_id we retrieved
        max_id = oldest_tweet['id']         

        print "[*] Retrieved: %d Tweets (max_id: %d)" % (len(full_tweet_list),max_id)
    
        # sleep to handle rate limiting
        time.sleep(3)
        
        # send next request with max_id set
        tweet_list = download_tweets(username,200,max_id-1)
    
        # grab the oldest Tweet
        if len(tweet_list):
            oldest_tweet = tweet_list[-1]
        

    # add the last few Tweets
    full_tweet_list.extend(tweet_list)
        
    # return the full Tweet list
    return full_tweet_list 
    
#getting the target account   
question = raw_input("Whose tweets do you want to download (Don't include the @ symbol)?\n\n")
full_tweet_list = download_all_tweets(question)

#saving all to new Excel spreadsheet
dest_filename = (dt + " " + question + " tweets.xlsx")
headers = ["Created at", "ID", "Text", "Retweet?", "Favourited", "Retweeted", "Reply to tweet ID", "Repy to user ID", "Reply to screen name", "User screen name", "User followers"]
wwb = Workbook()
ws1 = wwb.active
ws1.title = question
ws1.append(headers)

# extract stuff from each tweet for saving to Excel
for tweet in full_tweet_list:
    
    created_at = tweet ["created_at"]
    identity = tweet ["id_str"]
    
    fav_count = tweet ["favorite_count"]
    
    status_reply = tweet ["in_reply_to_status_id_str"]
    user_reply = tweet ["in_reply_to_user_id_str"]
    screen_name_reply = tweet ["in_reply_to_screen_name"]    

    retweet_count = tweet["retweet_count"]
    
    retweeted = tweet["retweeted"]
    
    text = tweet["text"]
    
    user_followers_count = tweet["user"]["followers_count"]

    user_screen_name = tweet["user"]["screen_name"]

    # write the results to spreadsheet
    
    data = [created_at, identity, text, retweeted, fav_count, retweet_count, status_reply, user_reply, screen_name_reply, user_screen_name, user_followers_count]    
    ws1.append(data)

wwb.save(filename = dest_filename)

print "Done"
