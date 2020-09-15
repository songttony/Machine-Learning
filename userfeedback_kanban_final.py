# initialize the environment and import different useful modules
import csv
import random
import numpy as np
import string
    
# OS module used for file operations
import os
import sys

# keep the contents into files,e.g '/tmp/feedbacks.txt', for future use
def info_store(para_info,para_file):
    with open(para_file,"w") as f:
        f.write(para_info)
        f.close()
    #use append mode to store all the feedbackclean info
    with open(para_file[:-4]+'All.txt',"a+") as ft:
        ft.write(para_info)
        ft.close()

# tkinter module used for GUI
import tkinter
import tkinter.messagebox
from tkinter import *

# define main function
def main():
    '''
        #MML command:
        feedbackcleanfile = info_input('Cleaned Feedbacks Location')
        feedbackfile = info_input('Raw Feedbacks Location')
        feedbackfreq = info_input('Feedback Words Frequency')
    '''
    feedbackcleanfile = '/tmp/feedbacks_clean.txt'
    feedbackfile = '/tmp/feedbacks.txt'
    feedbackfreq = '/tmp/feedbacks_freq.txt'
    feedbacklink = '/tmp/feedbacks_link.txt'
    stopfilepath = info_input('Stop Words Location')
    
    htmltype = ''
    #while loop till the correct URL entered
    while True:
        htmlurl = info_input('Web URL')
        htmlsoup,htmltype = grab_html(htmlurl)
        if htmltype !='':
            break

    recur_main(htmltype,htmlsoup,htmlurl,feedbackfile,feedbackcleanfile,feedbackfreq,stopfilepath,feedbacklink)

#recurrsion function to handle the recurrsion links in each pages
def recur_main(para_htype,para_hsoup,para_hurl,para_fbfile,para_fbcfile,para_fbcfreq, para_stopfile,para_fblink):
    
    if para_htype == 'Link':
        linkslist = switch_case(para_htype,para_hsoup,para_fblink)
        while True:
            for htmlurl in linkslist:
                htmlsoup,htmltype = grab_html(htmlurl)
                #invalid URL(empty or duplicated with main entry), move to next one
                if htmltype == '' or htmlurl == para_hurl:
                    continue
                if htmltype == 'Link':
                    recur_main(htmltype,htmlsoup,htmlurl,para_fbfile,para_fbcfile,para_fbcfreq, para_stopfile,para_fblink)
                switch_case(htmltype,htmlsoup,para_fbfile)
                jieba_split(para_fbfile, para_stopfile, para_fbcfile)
                wordfreq_gen(para_fbcfile,para_fbcfreq,htmlurl)
                wordcloud_gen(para_fbcfile,htmlurl,para_stopfile)
            break
        tkinter.messagebox.showinfo("message","All links parsered")
    else:
        switch_case(para_htype,para_hsoup,para_fbfile)
        jieba_split(para_fbfile, para_stopfile, para_fbcfile)
        wordfreq_gen(para_fbcfile,para_fbcfreq,para_hurl)
        wordcloud_gen(para_fbcfile,para_hurl,para_stopfile)

# define 'switch - case' selection for auto grabbing in Python
def switch_case(para_htype, para_soup, para_fbfile):
    infoType = {
      "Table": table_excerption,
      "Text": text_excerption,
      "Link": links_excerption
    }
    func = infoType.get(para_htype,'Link')
    return func(para_soup,para_fbfile)

# GUI input text module to get URL, storage files, etc, to be grapped
def info_input(para_input):

    """
    # manul input the target URL link
    # urlpage = input('Pls input the URL to be grapped:') 
    """
    # Define termination function 
    def return_callback(event):
        topwindow.quit()

    # Define close function for no input
    def close_callback():
        tkinter.messagebox.showinfo('message', 'No Input')
        # Terminate process
        topwindow.destroy()
        sys.exit()

    topwindow = tkinter.Tk()
    topwindow.title("Feedbacks Collection")

    url_label = tkinter.Label(topwindow, text = para_input)
    url_label.pack(side = LEFT)

    url_entry = tkinter.Entry(topwindow, bd=5)
    url_entry.pack(side = RIGHT)

    # Bind "Return" keyinput with Entry box
    url_entry.bind('<Return>', return_callback)

    # Define window close event
    topwindow.protocol("WM_DELETE_WINDOW", close_callback)

    topwindow.mainloop()

    # Get window input on URL to be grapped
    url_page = tkinter.Entry.get(url_entry)

    #topwindow.destroy()
    #hide the input window
    #used specially for while-loop in case wrong value input
    topwindow.withdraw()

    return url_page

# request html contents and store the HTML source data into variable: page
# redefine the default Python Header to avoid those Web with anti-crawler setting. 
# key is the setting of "User-Agent"
def grab_html(para_urlpage):
    
    #Grap the Web text data by using different WebScrap module
    """
    #requests module, used to get the target URL data
    import requests
    """
    htmltype = ''

    #Webscrap module: BeautifulSoup
    #pip3 install BeautifulSoup to install beautifulsoup package into python environment
    from bs4 import BeautifulSoup

    # WebScrap module: urllib
    import urllib.request 
    
    headers = {
        'Accept': '*/*',
        'Accept-Language': 'en-US,en;q=0.8',
        'Cache-Control': 'max-age=0',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36',
        'Connection': 'keep-alive'
        }

    try:
        #Change the default Python Headers
        urlpage_Head = urllib.request.Request(para_urlpage,None,headers=headers)
        """
        #another method:
        #use urllib parsing and store formated data into html
        #import urllib.parse
        #html = page.read()
        #html = html.decode('utf-8')
        """
        #Grap the page contents
        page = urllib.request.urlopen(urlpage_Head)
        
        # Use Beautiful Soup to parse HTML data and store in "soup"
        soup = BeautifulSoup(page,'html.parser')

        #check the contents is Table only and return the contents type for automatic grabbing data
        if soup.find('table', attrs={'class': 'tableSorter2'}) == None:
            try:
                soup.find('div',attrs={'class':'articleList'}).find_all('a')
                htmltype ='Link'
            except BaseException:
                htmltype = 'Text'
        else:
            htmltype = 'Table'

    except BaseException:
        #tkinter.messagebox.showinfo('Message',"Invalid URL")
        htmltype = ''
        soup = ''

    return soup,htmltype

# grab the table comments data and store into the file
# example used: https://www.fasttrack.co.uk/league-tables/tech-track-100/league-table/
def table_excerption (para_soup, para_fbfile):

    from bs4 import BeautifulSoup
    # WebScrap module: urllib
    import urllib.request

    def tableData(para_htmlinfo):
        UserFeedBack = []
        comments = ""
        for result in para_htmlinfo:
            # find every "td" contents in a particular row of table
            UserFeedBack = result.find_all('td')
            # if null, skip this row of the table
            if len(UserFeedBack) == 0: 
                continue
            else:
                # Store the found contents into the string variable and split into different line
                comments = comments + '\n' + UserFeedBack[7].getText()
                rank = UserFeedBack[0].getText()
                company = UserFeedBack[1].getText()
                location = UserFeedBack[2].getText()
                yearend = UserFeedBack[3].getText()
                salesrise = UserFeedBack[4].getText()
                sales = UserFeedBack[5].getText()
                staff = UserFeedBack[6].getText()  
        return comments

    #check abnormal situation in case of empty data grapped from HTML
    try:
        # find the table information. use the html page, right click "inspect" and locate the necessary 
        # tag and info where they are to be processed
        table = para_soup.find('table', attrs={'class': 'tableSorter2'})
        results = table.find_all('tr')
        contents = tableData(results)
        info_store(contents,para_fbfile)
    except BaseException:
        tkinter.messagebox.showinfo("message","No data parsered")
        sys.exit()

    return contents

# grab text data and store into the file
# example used: http://blog.sina.com.cn/s/blog_4701280b0101854o.html
def text_excerption (para_soup,para_fbfile):
    
    from bs4 import BeautifulSoup
    # WebScrap module: urllib
    import urllib.request 

    #check abnormal situation in case of empty data grapped from HTML
    try:
        #find the text contents
        context = para_soup.find('div', attrs={'class':'articalContent'})

        #error handling: if not 'div/class:articalContent' mode, get all the text;
        #otherwise, get the right contents
        if context is None:
            contents = para_soup.getText()
        else:
            #Grap the text contents by using build-in getText function
            contents = context.getText()

        info_store(contents,para_fbfile)
    except BaseException:
        tkinter.messagebox.showinfo("message","No data parsered")
        sys.exit()

    return contents

# grab URL data and store into the file
# example used: http://blog.sina.com.cn/s/articlelist_1191258123_0_1.html 
def links_excerption (para_soup, para_fblink):
    
    from bs4 import BeautifulSoup
    # WebScrap module: urllib
    import urllib.request

    #check abnormal situation in case of empty data grapped from HTML
    try:
        #find the url list
        urllists = para_soup.find('div',attrs={'class':'articleList'}).find_all('a')
        link_info = ''
        link_list = {}
        for link in urllists:
            link_info = link_info+'\n'+link.get('href')
        #Use Set's attribute to convert string(link_info) into set(link_list) so as to remove duplication
        link_list = set(link_info.splitlines())
        #convert set back to string by using join function so as to store into a file
        link_info = '\n'.join(link_list)
        info_store(link_info,para_fblink)
    except BaseException:
        tkinter.messagebox.showinfo("message","No data parsered")
        sys.exit()
    return link_list

# use jieba module to split the sentence into words
def jieba_split(para_fbfile, para_stopfile, para_fbcfile):
    
    #import jieba, the useful Chinese words split python module
    #pip3 install jieba to install jieba package into the python environment
    import jieba
    # Use the opensource stopwords DB # pip3 install stop_words to install
    from stop_words import get_stop_words

    # define stop words function
    def stopwordslist(para_filepath):
        try:
            stopwords = [line.strip() for line in open(para_filepath, 'r').readlines()]
        except BaseException:
            # Or get English stopwords from opensource stop_words DB
            stopwords = get_stop_words('english')

        return stopwords
 
    # define words cut function & remove those words in stop list
    def seg_sentence(para_sentence):
        sentence_seged = jieba.cut(para_sentence,cut_all=True)
        # Only keep the alphabet and remove the number and non-alphabet
        sentence_segedAlpha = filter(str.isalpha, sentence_seged)
        # specify stopwords file path (self customized stopwords list)
        stopwords = stopwordslist(para_stopfile)

        outstr = ''
        for word in sentence_segedAlpha:
            # Remove stopwords and sigle alphabet
            if word.lower() not in stopwords and len(word)>1:
                if word != '\t':
                    outstr += word
                    outstr += " "
        return outstr
 
    # Load the source file
    Feedbacks =open(para_fbfile,'r').read()

    # Data clean up
    wordlist = seg_sentence(Feedbacks)

    # Put cleaned data into file for storage
    #FeedBacksClean = open(feedbackcleanfile, 'w')
    #FeedBacksClean.write(wordlist)
    #FeedBacksClean.close()
    info_store(wordlist,para_fbcfile)

# count the frequency of the words
def wordfreq_gen(para_fbcfile,para_fbcfreq,para_link):
    
    import jieba
    from collections import Counter
    import operator

    # WordCount
    with open(para_fbcfile, 'r') as fr:
        data = jieba.cut(fr.read())
    data = dict(Counter(data))
 
    # Store words counter info into the file
    with open(para_fbcfreq, 'w') as fw:
        for k,v in sorted(data.items(),key=operator.itemgetter(1),reverse=True):
            if k !=" ":     #Remove " " value
               fw.write('%s,%d\n' % (k, v))
               info_store('%s,%d\n' % (k,v),para_fbcfreq)

    #import wordcloud for generate word frequency dashboard
    #pip3 install wordcloud to install wordcloud into the python environment

    import matplotlib.pyplot as plt

    #change fonts support Chinese # Solve matplotlib doesn't support Chinese issue
    from pylab import mpl
    mpl.rcParams['font.sans-serif'] = ['FangSong'] # 指定默认字体
    mpl.rcParams['axes.unicode_minus'] = False # 解决保存图像是负号'-'显示为方块的问题

    #generate the histogram of word frequency
    #use sorted function to sort the data dictionary. key = operator.itemgetter(1) 
    #or key = lambda item:item[1]
    for k,v in sorted(data.items(),key=operator.itemgetter(1),reverse=True):
        if k !=" " and v > 3:     #Don't display " " and frequency <=3
            #change fontsize
            plt.tick_params(axis='x', labelsize=4)
            #Rotate -15 degree
            plt.xticks(rotation = -15)
            plt.bar(k,v)
    plt.title("Feedback Hot Words"+'\n'+para_link)
    #set y-axis' scale 0 -100
    plt.ylim(0, 100)
    plt.ylabel("Count#")
    plt.xlabel("Feedbacks")
    plt.show()    

# Generate Wordcloud
def wordcloud_gen(para_fbcfile,para_link,para_stopfile):
    from wordcloud import WordCloud
    import matplotlib.pyplot as plt
    import operator
    FeedBacksClean =open(para_fbcfile,'r').read()
    #set Chinese fonts to sovle the Chinese display issue
    #default fontpath in wordcloud doesn't support Chinese display
    fontpath =r'C:/library/fonts/Microsoft/SimHei.ttf'
    if FeedBacksClean !='':
        UserfeedbackCount = WordCloud(font_path=fontpath,stopwords=para_stopfile).generate(FeedBacksClean)
        plt.imshow(UserfeedbackCount)
        plt.title('Word Cloud'+'\n'+para_link)
        plt.axis("off")
        plt.show()

if __name__ == '__main__':
    #Specify the file location
    main()