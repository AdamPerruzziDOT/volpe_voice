#This script:
#   -Reads HTML from V120 blogposts
#   -Identifies links to Volpe Project Dashboards
#   -Writes the links and several other attributes to a text file
#
#Example links of interest are as follows:
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/Volpe-Center-AllInOne.aspx
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/Tech-Center-All.aspx?TechCenter=V-310
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/Division-All.aspx?Division=V-311
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/ProjectMaster-All.aspx?Project=HW9G
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/Project-All.aspx?Project=HW9GA200
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/Sponsor-All.aspx?Sponsor=AIR FORCE
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/Portfolio-All.aspx?Portfolio=DHS
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/Deliverable.aspx?ProjectLibraryDocId=HW9GA100%7C31786
#   -http://spminiapps.volpe.dot.gov/sites/DW/Pages/Staff.aspx?InputName=Richardson,%20Heather
#
#Input information is:
#   -Login information [config.txt]
#
#Output files are:
#   -Article links to be placed on the dashboards [volpe_voice_dash_links_YYYYMMDD.txt]
#   -Errors file, to be corrected [errors_YYYYMMDD.xlsx]
#
#Script produced by:
#   -Alex Linthicum, USDOT Volpe Center
#   -Adam Perruzzi, USDOT Volpe Center
#
#Originally posted to GitHub on 2017 July 28th



###Libraries
import math
import os
import pandas as pd
import re
import requests
import shutil
import sys
import time
import unidecode
from bs4 import BeautifulSoup
from nltk.tokenize import sent_tokenize
from nltk.tokenize import word_tokenize
from requests_ntlm import HttpNtlmAuth



###Tests whether links on a page are to the Dashboards
def is_dash_link(href):
    return href and re.compile('DW\/Pages').search(href) #Links with 'DW/Pages' in the link address


###Cleans up the text to remove or replace undesirable characters
def cleanUnicode(text):
    text = text.replace(u'\u00a0',' ') #No-break space
    text = text.replace(u'\u200b','') #Zero-width space
    text = text.replace(u'\u2018','\'') #Left single quote
    text = text.replace(u'\u2019','\'') #Right single quote
    text = text.replace(u'\u00A0',' ') #No-break space for uppercase
    text = text.replace(u'\u200B','') #Zero-width space for uppercase
    text = text.replace(u'\u2013','-') #En-dash
    text = text.replace(u'\u2014','-') #Em-dash
    text = text.replace(u'\u201c','\"') #Double-quote
    text = text.replace(u'\u201d','\"') #Double-quote
    text = text.replace(u'\u2026','...') #Ellipsis
    text = text.replace(u'\u200e','') #Left to right mark
    return text


###Rebuilds a sentence from an array of tokens, except for line breaks
def untokenize(words): 
    text = ' '.join(words)
    step1 = text.replace("`` ", '"').replace(" ''", '"').replace('. . .',  '...')
    step2 = step1.replace(" ( ", " (").replace(" ) ", ") ")
    step3 = re.sub(r' ([.,:;?!%]+)([ \'"`])', r"\1\2", step2)
    step4 = re.sub(r' ([.,:;?!%]+)$', r"\1", step3)
    step5 = step4.replace(" '", "'").replace(" n't", "n't").replace("can not", "cannot")
    step6 = step5.replace(" ` ", " '")
    return step6.strip()


###Indicates whether or not a dashboard item is linked properly
def properCategory(link,categories):
    category = link.split('DW/Pages/')[-1].split('.')[0] #How the link was actually categorized [e.g. division, staff, etc.]
    target = link.split('=')[-1] #What the link leads to [e.g. a sponsor, division, etc.]
    if target.lower() in categories: #If the target has a proper category
        if categories[target.lower()] == category.lower(): #If the target is properly categorized
            return [True,'',cleanCategory(category.lower()),target] #Indicate the item is properly linked
        else: #If the target is not properly categorized
            correct = 'http://spminiapps.volpe.dot.gov/sites/DW/Pages/' + categories[target.lower()] + '.aspx?' #Dashboard link base, including the correct category
            if categories[target.lower()] == 'tech-center-all': #If the proper category is a technical center
                correct += 'TechCenter=' #Add the correct entity label
            elif categories[target.lower()] == 'division-all': #If the proper category is a divison
                correct += 'Division=' #Add the correct entity label
            elif categories[target.lower()] == 'toplevel': #If the proper category is a top level organization
                correct += 'Org=' #Add the correct entity label
            elif categories[target.lower()] == 'operations': #If the proper category is an operations organization
                correct += 'Org=' #Add the correct entity label
            elif categories[target.lower()] == 'sponsor-all': #If the proper category is a sponsor
                correct += 'Sponsor=' #Add the correct entity label
            correct += target #Add the actual dashboard item to the end of the correct link
            return [False,correct,cleanCategory(category.lower()),target] #Indicate the item is not properly linked, return the correct version
    elif category not in ['Project-all','Staff']: #Link points to an unrecognized object
        return [False,'',cleanCategory(category.lower()),target] #Indicate the item is not properly linked, and there is no available correction
    else: #Link points to Project or Staff
        return [True,'',cleanCategory(category.lower()),target]


###Returns the proper category name, based on the link category name
def cleanCategory(category):
    if category == 'tech-center-all':
        return 'Tech Center'
    elif category == 'division-all':
        return 'Division'
    elif category == 'toplevel':
        return 'Top Level'
    elif category == 'operations':
        return 'Operations'
    elif category == 'sponsor-all':
        return 'Sponsor'
    elif category == 'project-all':
        return 'Project'
    elif category == 'staff':
        return 'Staff'
    else:
        return 'UNK'



if __name__ == '__main__':
    
    ###Initialize Web Session
    cfgFile = open('config.txt','r') #Open config file
    cfgInfo = cfgFile.readlines() #Lines of file to array
    cfgFile.close() #Close config file
    username = 'ADDOT\\' + cfgInfo[0].split('Username:')[1].strip() #Retrieve username, prefixed with ADDOT domain
    password = cfgInfo[1].split('Password:')[1].strip() #Retrieve password
    s = requests.Session() #Create webserver session
    s.auth = HttpNtlmAuth(username,password) #Authenticate
    
    
    ###Determine starting place, based on last file
    recentFileDate = 0 #YYYYMMDD date as number
    for file in os.listdir(sys.path[0]): #Each file in the same folder as this script
        if '.txt' in file and 'volpe_voice_dash_links' in file: #If this is one of the link log files
            fileDate = file.split('.')[0] #Remove file extension
            fileDate = fileDate.split('_')[-1] #Extract date string only
            if int(fileDate) > recentFileDate: #If the current file is newer than all previous files
                recentFileDate = int(fileDate) #Update the newest date
    
    recentFileName = 'volpe_voice_dash_links_'+str(recentFileDate)+'.txt' #Full filename, based on newest date
    recentFile = open(recentFileName,'r') #Open most recent links file
    recentFileInfo = recentFile.readlines() #Lines of file to array
    recentFile.close() #Close most recent links file
    startPage = int(recentFileInfo[-1].split('|')[4].split('=')[-1])+1 #Extract most recent article number, and add one
    endPage = startPage + 25 #Set end page 25 pages ahead of starting page
    
    
    ###Identify pages that exist, to be scraped
    print('Starting at page: '+str(startPage)) #Alert the user of starting place
    volpePostIDs = [] #Array to hold numbers of all pages that exist
    x = startPage #Start one ahead of the most recent article
    while x <= endPage: #Until there are 50 pages in a row that don't exist
        url_str = 'http://spmain.volpe.dot.gov/InternalNews/lists/posts/VolpePost.aspx?ID=' + str(x) #Create link to check
        r = s.head(url_str) #Retreieve page, using persisting session
        if r.status_code < 400 and 'SharePointError' not in r.headers: #If the page exists
            print(x) #Log the page number for the user
            volpePostIDs.append(x) #Add the page to the list
            endPage = x + 25 #Always checking 25 pages after the last found article
        x+= 1 #Advance to next page
    print('Completed identification of ' + str(len(volpePostIDs)) + ' pages') #Alert user of total number of articles found
    
    
    ###Setup for page scan
    print('Extracting links...') #Alert the user that links are being extracted
    errors = [] #Running list of pages with errors, to be addressed manually
    categories = {} #Dictionary for checking proper link category
    for line in cfgInfo[2:]: #For all config file lines after the second
        for target in line.split(':\t')[1].strip().split(', '): #For each group member in the list
            categories[target.lower()] = line.split(':\t')[0].lower() #Create dictionary entry as [member] = group
    linkSkip = ['http://spminiapps.volpe.dot.gov/sites/DW/Pages/Volpe-Center-AllInOne.aspx', 'http://spminiapps.volpe.dot.gov/sites/DW/Pages/Home.aspx'] #Links to skip
    concMin = 25 #Minimum words in a concordance
    concMax = 30 #Maximum words in a concordance
    str_print = '' #String to be written to output file at the end of link collection
    
    
    ###Scan pages for links
    for num in volpePostIDs: #For each article that was found
        
        
        ###General page information
        print('Page ' + str(num) +'...') #Log article number for the user
        url_str = 'http://spmain.volpe.dot.gov/InternalNews/lists/posts/VolpePost.aspx?ID=' + str(num) #Link to page
        r = s.get(url_str) #Get page content, using persisting session
        soup = BeautifulSoup(r.text, "html.parser") #Parse the page text using BeautifulSoup
        bpTitle = unidecode.unidecode(soup.find_all('h3', class_="blogPostTitle")[0].string).strip() #Article title
        bpDate = unidecode.unidecode(soup.find_all('h4', class_="blogPostDate")[0].string).strip() #Post data
        
        
        ###Clean up the page text
        pageSent = [] #List of sentences in the article
        for td in soup.find_all('td', class_='ms-vb blogPost'): #Page content table cell
            for string in td.stripped_strings: #Each string within the table cell, with whitespace removed
                string = unidecode.unidecode(string) #Clean up the text
                string = string.replace('\n',' ').replace('\r',' ').strip() #Remove both types of newlines and any whitespace
                pageSent.extend(sent_tokenize(string)) #Add the sentences in this string to the list of article sentences
        if pageSent[-1][:6].lower() == 'posted': #If the final sentence is the posting information
            pageSent = pageSent[:-1] #Remove the last sentence
        
        
        ###Process each dashboard link on the page
        for link in soup.find_all(href=is_dash_link): #For each dashboard link on the page
            if link.get('href') not in linkSkip and cleanUnicode(link.text).replace('\n','').replace('\r','').strip() not in ['',',']: #No empty, comma, or skipped links
                
                
                ###Check proper categorization
                categoryEval = properCategory(link.get('href'),categories) #Retrieve categorization status of the link, along with any corrections
                if not categoryEval[0]: #If the link was not properly categorized
                    errors.append({'Page Number': num, 'Link': url_str, 'Type': 'Link', 'Problem': link.get('href'), 'Correction': categoryEval[1]}) #Store in error list
                
                
                ###Get search term
                success = True #Was the link able to be successfully extracted?
                searchTerm = unidecode.unidecode(link.text) #Retrieve and clean up the link text
                searchTerm = searchTerm.replace('\n',' ').replace('\r',' ').strip() #Remove line breaks and whitespace
                divSearch = re.search('V-[0-9][0-9][0-9]',searchTerm) #Searching for a 'V-###' pattern
                divSearchMod = re.search('[0-9][0-9][0-9]',searchTerm) #Searching for a '###' pattern
                if divSearch: #If the search term is a division
                    searchTerm = divSearch.group() #Extract just 'V-###'
                elif divSearchMod: #If it's a malformed division link
                    searchTerm = 'V-' + str(divSearchMod.group()) #Add the 'V-' front to the numbers-only term
                else: #If the link isn't to a division
                    while True: #Until the link is clean or 'dissolved'
                        if searchTerm[-1].isalpha(): #If the last character is alpha
                            if searchTerm[-2:] == '\'s': #If the term ends with a posessive
                                searchTerm = searchTerm[:-2] #Remove the posessive
                            break #The link is clean; exit the while loop
                        else: #If the last term is not alpha
                            if len(searchTerm) > 1: #If the string is longer than one character
                                searchTerm = searchTerm[:-1] #Shorten the term by one character
                            else: #If the string is one or fewer characters long
                                searchTerm = unidecode.unidecode(link.text).replace('\n','').replace('\r','').strip() #Retrieve original search term
                                errors.append({'Page Number': num, 'Link': url_str, 'Type': 'Search Term', 'Problem': searchTerm, 'Correction': ''}) #Store in error list
                                success = False #Indicate the link was not successfully extracted
                                break #The link is dissolved; exit the while loop
                searchTerm = searchTerm.strip() #Remove any additional whitespace
                searchTerm = re.sub(' +',' ',searchTerm) #Condense blocks of multiple spaces
                print('<' + searchTerm + '>') #Log the search term to the console, for the user
                
                
                ###Process link
                if success: #If the search term was able to be found in the previous step
                    
                    
                    ###Get concordance
                    concord = '' #String that will eventually become the surrounding summary text
                    for i in range(0,len(pageSent)): #For each of the sentences on the page
                        
                        ###Generate text to pull concordance from
                        if searchTerm in pageSent[i]: #If the current sentence contains the search term
                            concord = pageSent[i] #Initialize the concordance as the sentence the search term is in
                            j = 0 #Number of sentences ahead of the matching sentence, in the page text
                            k = 0 #Number of sentences behind the matching sentence, in the page text
                            while len(word_tokenize(concord)) < concMin: #Until there are enough words in the concordance
                                if (i+j+1) < len(pageSent): #If the current sentence is not the last on the page
                                    j=j+1 #Move one sentence ahead in the page text
                                    concord = concord + ' ' + pageSent[i+j] #Add the next sentence to the concordance
                                elif (i-k) > 0 : #Last sentence already included, not first sentence, add previous
                                    k=k+1 #Move one sentence back in the page text
                                    concord = pageSent[i-k] + ' ' + concord #Add the previous sentence to the concordance
                                else: #Couldn't find any additional sentences to add
                                    break #Stop searching for additional sentences, and accept the current concordance
                            concord_W = word_tokenize(concord) #Make a list of words to draw concordance from
                            
                            
                            ###Shorten the concordance to the appropriate length, if necessary
                            if len(concord_W) > concMax: #If the list of words is too long for a concordance
                                if k == 0 & j != 0: #Only went forwards
                                    concord_W = concord_W[:30] #Select the first 30 words
                                    concord = untokenize(concord_W) + '...' #Recombine list of words, add ellipsis
                                elif k !=0: #May have extended both ways, break in the front
                                    concord_W = concord_W[(len(concord_W)-29):] #Select at most 30 words
                                    concord = '...' + untokenize(concord_W) #Recombine, add ellipsis to the front
                                else: #A single large sentence
                                    searchTerm_W = word_tokenize(searchTerm) #Get words of search term
                                    searchTerm_F = searchTerm_W[0] #Front word in search term
                                    searchTerm_R = searchTerm_W[len(searchTerm_W)-1] #Last word in search term, even if same word
                                    
                                    
                                    ###Locate the search term within the sentence
                                    index_F = 0 #Position of front word
                                    index_R = 0 #Position of rear word
                                    for word in range(0,len(concord_W)-1): #For each word in the single concordance sentence
                                        if concord_W[word] == searchTerm_F: #If the front word of the search term has been found
                                            index_F = word #Stores index of the front word of the search term
                                        if concord_W[word] == searchTerm_R: #If the rear word of the search term has been found
                                            index_R = word #Stores index of the back of the match
                                            break #Stop searching the sentence
                                    
                                    
                                    ###Center the concordance on the search term
                                    index_range = index_R-index_F #Difference between front and rear word. Gives length to subtract
                                    index_F = index_F - math.floor((concMax-index_range)/2) #Extend front end by half the remaining concordance length
                                    index_R = index_R + math.floor((concMax-index_range)/2) #Extend rear end by half the remaining concordance length
                                    if index_F < 0: #If there were not enough words in front of the search term
                                        index_R = index_R - index_F #Add more words to the end, to replace the missing front ones
                                        index_F = 0 #Set the start of the concordance to the start of the sentence
                                    if index_R >= len(concord_W): #If there were not enough words after the search term
                                        index_F = index_F - (index_R - len(concord_W)+1) #Add more words to the front, to replace the missing end ones
                                        index_R = len(concord_W)-1 #Set the end of the concordance to the end of the sentence
                                    if index_F < 0: #If a second addition moved the beginning of the concordance too far
                                        index_F = 0 #Set the start of the concordance to the start of the secntence
                                    
                                    
                                    ###Set final concordance
                                    oldMax = len(concord_W)-1 #number of words in the sentence before slicing
                                    concord_Final = concord_W[int(index_F):int(index_R)] #Extract the final concordance
                                    concord = untokenize(concord_Final) #Recombine the tokens
                                    if index_F > 0: #If the concordance started in the middle of a sentence
                                        concord = '...' + concord #Add an ellipsis to the front
                                    if index_R < oldMax: #If the concordance ended in the middle of a sentence
                                        concord = concord + '...' #Add an ellipsis to the end
                                concord = concord.replace('....','...') #Shorten final ellipsis if it occurs after a period
                            break #Exit, once the sentence containing the search term has been found
                        elif i == (len(pageSent)-1): #Did not find the search term in any sentence
                            errors.append({'Page Number': num, 'Link': url_str, 'Type': 'Concordance', 'Problem': searchTerm, 'Correction': ''}) #Store in list of errors
                    
                    ###Add all fields of interest to print string
                    str_print += str(categoryEval[2]) + '|' + str(categoryEval[3]) #Add category information to print string
                    str_print += '|"' + str(bpTitle) +'"|'+ str(bpDate) +'|'+ str(url_str) #Add post information to print string
                    str_print += '|"'+ str(concord) +'"\n' #Add concordance information to print string
    
    
    ###Write output files
    print('Scan complete. Writing files...') #Notify the user the output phase has begun
    
    
    ###Relocate the errors file, regardless
    for file in os.listdir(sys.path[0]): #Each file in the same folder as this script
            if file == 'volpe_voice_errors.xlsx': #If there is an existing error file
                while True: #Loop until the errors file has been successfully relocated
                    try: #If the file is accessible
                        shutil.move(os.path.join(sys.path[0],file),os.path.join(sys.path[0],'Old Error Logs','volpe_voice_errors_' + time.strftime('%Y%m%d_%H%M%S') + '.xlsx')) #Relocate the old errors file
                        break #Exit the loop, once the move is completed
                    except: #If the file is inaccesible (likely open)
                        placeholder = input('Please close error file. Press [Enter] when ready...') #Give the user time to close the error file, then advance
    
    
    ###Print errors file or link file, depending on the success of the script
    if not errors: #If there were no errors on any of the scanned pages
        while True: #Loop until the link file has been successfully relocated
            try: #If the file is accessible
                shutil.copyfile(os.path.join(sys.path[0],recentFileName),os.path.join(sys.path[0],'Old Link Files',recentFileName)) #Create a backup of the old link file
                os.rename(os.path.join(sys.path[0],recentFileName),os.path.join(sys.path[0],'volpe_voice_dash_links_' + time.strftime('%Y%m%d') + '.txt')) #Rename using the date
                break #Exit the loop, once the move is completed
            except: #If the file is inaccesible (likely open)
                placeholder = input('Please close link file. Press [Enter] when ready...') #Give the user time to close the link file, then advance
        outFile = open('volpe_voice_dash_links_' + time.strftime('%Y%m%d') + '.txt', 'a', encoding='utf8') #Open the recently copied file to append to
        if str_print != '': #If the string to print isn't empty
            outFile.write('\n' + str_print[:-1]) #Add a new line to advance from old entries, and add the print string, minus the last line return
        outFile.close() #Close the output file
    else: #There were errors in some of the pages being checked
        df = pd.DataFrame(errors) #Put the errors into a dataframe for exporting
        df = df[['Page Number','Link','Type','Problem','Correction']] #Re-order the columns
        writer = pd.ExcelWriter('volpe_voice_errors.xlsx') #Name of the workbook to be written to
        df.to_excel(writer,sheet_name='Errors') #Sheet to write the dataframe to
        writer.save() #Close the output workbook