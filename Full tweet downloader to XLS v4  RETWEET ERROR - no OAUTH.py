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
#headers = ["Created at", "ID", "Text", "Retweet?", "Favourited", "Retweeted", "Lang", "Source", "Coords", "Place", "Reply to tweet ID", "Repy to user ID", "Reply to screen name", "Hashtags", "URLs", "User name", "User screen name", "User ID", "User created at", "User description", "User favourites", "User followers", "User friends", "User lang", "User lists count", "User location", "User tweets", "User time zone", "User URL", "User verified?", "User mentions", "RT created at", "RT ID", "RT Text", "RT Favourited", "RT Retweeted", "RT Lang", "RT Source", "RT Coords", "RT Place", "RT Reply to tweet ID", "RT Repy to user ID", "RT Reply to screen name", "RT Hashtags", "RT URLs", "RT user mentions", "RT User name", "RT User screen name", "RT User ID", "RT User created at", "RT User description", "RT User favourites", "RT User followers", "RT User friends", "RT User lang", "RT User lists count", "RT User location", "RT User tweets", "RT User time zone", "RT User URL", "RT User verified?"]
headers = ["Created at", "ID", "Text", "Retweet?", "Favourited", "Retweeted", "Lang", "Source", "Coords", "Reply to tweet ID", "Repy to user ID", "Reply to screen name", "User name", "User screen name", "User ID", "User created at", "User description", "User favourites", "User followers", "User friends", "User lang", "User lists count", "User location", "User tweets", "User time zone", "User URL", "User verified?", "RT created at", "RT ID", "RT Text", "RT Favourited", "RT Retweeted", "RT Lang", "RT Source", "RT Coords", "RT Reply to tweet ID", "RT Repy to user ID", "RT Reply to screen name", "RT User name", "RT User screen name", "RT User ID", "RT User created at", "RT User description", "RT User favourites", "RT User followers", "RT User friends", "RT User lang", "RT User lists count", "RT User location", "RT User tweets", "RT User time zone", "RT User URL", "RT User verified?"]
wwb = Workbook()
ws1 = wwb.active
ws1.title = question
ws1.append(headers)

# extract stuff from each tweet for saving to Excel
for tweet in full_tweet_list:
    
    coords = tweet["coordinates"]
    created_at = tweet ["created_at"]
    identity = tweet ["id_str"]
    
    hashtags = []
    for hashtag in tweet ["entities"]["hashtags"]:
        hashtags.append (hashtag["text"])
    
    urls = []
    for url in tweet ["entities"]["urls"]:
        urls.append(url["expanded_url"])
        
    user_mentions = []
    for mention in tweet["entities"]["user_mentions"]:
        user_mentions.append (mention["screen_name"])

    fav_count = tweet ["favorite_count"]
    
    status_reply = tweet ["in_reply_to_status_id_str"]
    user_reply = tweet ["in_reply_to_user_id_str"]
    screen_name_reply = tweet ["in_reply_to_screen_name"]    

    tweet_lang = tweet["lang"]
    
    #try:
     #   tweet_place = tweet["place"]["full_name"]
    #except:
     #   pass
#did the above work?    
    
    retweet_count = tweet["retweet_count"]
    
    retweeted = tweet["retweeted"]
    
    

    if tweet["retweeted_status"]["coordinates"]:
        rt_coords = tweet["retweeted_status"]["coordinates"]
    else:
        rt_coords = None

    if tweet["retweeted_status"]["created_at"]:
        rt_created_at = tweet["retweeted_status"]["created_at"]
    else:
        rt_created_at = "N/A"
        

    rt_coords = tweet["retweeted_status"]["coordinates"]
    rt_created_at = tweet["retweeted_status"]["created_at"]
        
    rt_hashtags = []
    for rt_hashtag in tweet ["retweeted_status"]["entities"]["hashtags"]:
        rt_hashtags.append (rt_hashtag["text"])
        
    rt_urls = []
    for rt_url in tweet ["retweeted_status"]["entities"]["urls"]:
        rt_urls.append(rt_url["expanded_url"])        
            
        
    rt_user_mentions = []
    for rt_mention in tweet["retweeted_status"]["entities"]["user_mentions"]:
        rt_user_mentions.append (rt_mention["screen_name"])        
        
    rt_fav_count = tweet["retweeted_status"]["favorite_count"]
        
    rt_identity = tweet["retweeted_status"]["id_str"]
    
    rt_in_reply_to_screen_name = tweet["retweeted_status"]["in_reply_to_screen_name"]
    rt_in_reply_to_status_id = tweet["retweeted_status"]["in_reply_to_status_id_str"]
    rt_in_reply_to_user_id = tweet["retweeted_status"]["in_reply_to_user_id_str"]
        
    rt_lang = tweet["retweeted_status"]["lang"]
        #try:
         #   rt_place = tweet["retweeted_status"]["place"]["full_name"]
        #except:
         #   pass
#did the above work?         
        
    rt_retweet_count = tweet["retweeted_status"]["retweet_count"]
    rt_source = tweet["retweeted_status"]["source"]
    rt_text = tweet["retweeted_status"]["text"]
        
    rt_user_created_at = tweet["retweeted_status"]["user"]["created_at"]
    rt_user_description = tweet["retweeted_status"]["user"]["description"]

    rt_user_favourites_count = tweet["retweeted_status"]["user"]["favourites_count"]
    rt_user_followers_count = tweet["retweeted_status"]["user"]["followers_count"]
    rt_user_friends_count = tweet["retweeted_status"]["user"]["friends_count"]
    rt_user_id = tweet["retweeted_status"]["user"]["id_str"]
    rt_user_lang = tweet["retweeted_status"]["user"]["lang"]
    rt_user_listed_count = tweet["retweeted_status"]["user"]["listed_count"]
    rt_user_location = tweet["retweeted_status"]["user"]["location"]
    rt_user_name = tweet["retweeted_status"]["user"]["name"]
    rt_user_screen_name = tweet["retweeted_status"]["user"]["screen_name"]
    rt_user_statuses_count = tweet["retweeted_status"]["user"]["statuses_count"]
    rt_user_time_zone = tweet["retweeted_status"]["user"]["time_zone"]
    rt_user_url = tweet["retweeted_status"]["user"]["url"]
    rt_user_verified = tweet["retweeted_status"]["user"]["verified"]
    
    source = tweet["source"]
    text = tweet["text"]
    
    user_created_at = tweet["user"]["created_at"]
    user_description = tweet["user"]["description"]
    
    user_favourites_count = tweet["user"]["favourites_count"]
    user_followers_count = tweet["user"]["followers_count"]
    user_friends_count = tweet["user"]["friends_count"]
    user_id = tweet["user"]["id_str"]
    user_lang = tweet["user"]["lang"]
    user_listed_count = tweet["user"]["listed_count"]
    user_location = tweet["user"]["location"]
    user_name = tweet["user"]["name"]
    user_screen_name = tweet["user"]["screen_name"]
    user_statuses_count = tweet["user"]["statuses_count"]
    user_time_zone = tweet["user"]["time_zone"]
    user_url = tweet["user"]["url"]
    user_verified = tweet["user"]["verified"]

    # write the results to spreadsheet
    
#    data = [created_at, identity, text, retweeted, fav_count, retweet_count, tweet_lang, source, coords, tweet_place, status_reply, user_reply, screen_name_reply, hashtags, urls, user_name, user_screen_name, user_id, user_created_at, user_description, user_favourites_count, user_followers_count, user_friends_count, user_lang, user_listed_count, user_location, user_statuses_count, user_time_zone, user_url, user_verified, user_mentions, rt_created_at, rt_identity, rt_text, rt_fav_count, rt_retweet_count, rt_lang, rt_source, rt_coords, rt_place, rt_in_reply_to_status_id, rt_in_reply_to_user_id, rt_in_reply_to_screen_name, rt_hashtags, rt_urls, rt_user_mentions, rt_user_name, rt_user_screen_name, rt_user_id, rt_user_created_at, rt_user_description, rt_user_favourites_count, rt_user_followers_count, rt_user_friends_count, rt_user_lang, rt_user_listed_count, rt_user_location, rt_user_statuses_count, rt_user_time_zone, rt_user_url, rt_user_verified]
    data = [created_at, identity, text, retweeted, fav_count, retweet_count, tweet_lang, source, coords, status_reply, user_reply, screen_name_reply, user_name, user_screen_name, user_id, user_created_at, user_description, user_favourites_count, user_followers_count, user_friends_count, user_lang, user_listed_count, user_location, user_statuses_count, user_time_zone, user_url, user_verified, rt_created_at, rt_identity, rt_text, rt_fav_count, rt_retweet_count, rt_lang, rt_source, rt_coords, rt_in_reply_to_status_id, rt_in_reply_to_user_id, rt_in_reply_to_screen_name, rt_user_name, rt_user_screen_name, rt_user_id, rt_user_created_at, rt_user_description, rt_user_favourites_count, rt_user_followers_count, rt_user_friends_count, rt_user_lang, rt_user_listed_count, rt_user_location, rt_user_statuses_count, rt_user_time_zone, rt_user_url, rt_user_verified]    
    ws1.append(data)

wwb.save(filename = dest_filename)

print "Done"
