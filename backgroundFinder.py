import re
import nltk
import bs4
import excel
from bs4 import BeautifulSoup
from nltk import sent_tokenize
from nltk.corpus import stopwords

excel_path = "/Users/sameerhimati/Desktop/ISBStuff/BackgroundExcel.xlsx"
num_rows = excel.get_rows(excel_path,"Sheet1")
num_cols = excel.get_columns(excel_path,"Sheet1")
nltk.download("punkt")
nltk.download("stopwords")
stop_words = set(stopwords.words('english'))


def cleaner(file):
    """
    Uses beautiful soup the open and clean the text file.
    :param file:
    :return:
    """
    f = open(file, 'r')
    data = BeautifulSoup(f.read()).get_text()
    return data


def findNextNum(l, ind):
    """
    Finds all the numbers in the list l from the index ind.
    :param l:
    :param ind:
    :return:
    """
    buffer = l[ind:ind+500]
    loc = re.findall(r'\d\d', buffer)
    ilist = re.findall(r'[I][-]\d+', buffer)
    if len(ilist) != 0:
        number = re.findall(r'\d+', ilist[0])
        if number[0] == loc[0]:
            return loc
        location.insert(0, number[0])
    return loc


def clean(string):
    """
    Removes unnecessary new lines and white spaces.
    :param string:
    :return:
    """
    string = string.strip()
    string = string.replace('\n', ' ')
    string = string.replace('\t', ' ')
    string = re.sub(' +', ' ', string)
    return string

filenames = []
for row in range(2, 100):  # parse each row.
    if row == 51:  # this file is corrupt and gives a beautiful soup error.
        continue
    useful = []
    # below we extract the data from the excel.
    dealNumber = excel.read_data(excel_path, "Sheet1", row, 1)
    acquirer = excel.read_data(excel_path, "Sheet1", row, 7)
    target = excel.read_data(excel_path, "Sheet1", row, 8)
    threshold = excel.read_data(excel_path, "Sheet1", row, 9)
    threshold = int(threshold)  # this is the number of times background comes before the table of contents.
    txt = str(dealNumber)  # gets the deal number as a string.
    txt = txt + ".txt"  # txt extension added.
    #filenames.append(txt)
    data = cleaner(txt)  # opens the text file and cleans it using beautiful soup.
    sentenceList = nltk.sent_tokenize(data)  # uses nltk to tokenize the file into "Sentences".
    occurrence = 0  # a variable used to check the number of times we see "background" before the table of contents.
    ind = 0  # this is the index of where the title of the section is in the sentence list.
    wordLocation = 0  # this is the index of the title in the specific sentence.
    f = open("test.txt", "w")
    # this parses the sentence sentence list to see if we find the background section. Specifically in the table of
    # contents.
    for i in sentenceList:
        if "background" in i.lower():
            if occurrence < threshold+1:
                occurrence += 1
                if "background of the merger" in i.lower():  # background of the merger
                    ind = i.lower().rfind("background of the merger")  # background of the merger
                    wordLocation = sentenceList.index(i)
                elif "background" in i.lower():
                    ind = i.lower().rfind("background")
                    wordLocation = sentenceList.index(i)

    indexList = []
    indexList = findNextNum(sentenceList[wordLocation], ind)  # finds the page numbers of sections starting with the
    # background's
    print(indexList)
    if len(indexList) == 0:  # this is a precautionary check to make sure that there is no part of the text that we
        # missed. Since it is possible that the title occurs once after the first time we find the word "background".
        print("this")
        print(row)
        for i in sentenceList[wordLocation+1:wordLocation+30]:  # goes through all the sentences after the first time we
            # find background.
            testList = []
            testList = findNextNum(i, 0)  # finds the next number after the title.
            if len(testList) != 0:
                indexList.append(testList[0])  # if a number was found. Adds it. Twice.
                indexList.append(testList[0])
                break
    print(indexList)
    if len(indexList) == 0:  # after the check if it still doesn't have any page numbers.
        excel.write_data(excel_path, "Sheet1", row, 10, "No Background Section")
    else:
        location = sentenceList[wordLocation].find(indexList[0], ind)  # finds the index of the number in the line with
        # the title.
        title = sentenceList[wordLocation][ind:location]  # finds the title from the table of contents
        title = clean(title)  # clean to avoid needless errors in comparisons.
        bad_chars = [';', ':', '!', "*", "."]
        for i in bad_chars:  # removes special characters in the title extracted from the table of contents.
            title = title.replace(i, '')
        if title[-2:] == "I-":  # Few filings have page numbers as I-3, I-4 etc., this removes that potential error.
            title = title[:-2]
        print(title)
        title = clean(title)  # clean again in case.
        if len(indexList) == 1:  # rare case but in case the function only extracts one page number this makes sure
            # we extract that page.
            indexList.append(indexList[0])
        page1 = int(indexList[0])  # sets the background section's page number
        page2 = int(indexList[1])  # sets the next section's page number
        diff = (page2 - page1)  # this checks how big the section is. (page number difference.)

        if diff == 0:  # in case the background section is really small this ensures the rest of the code runs smoothly.
            page2 += 1
            diff += 1

        increment = 100*diff  # this is the increment that I set as an upper limit.
        potential = []  # this stores a list of lists of sentences.
        final = []  # used later. This contains one section from the list of potential sentences from the potential
        # variable above

        # this section below is particularly important as it finds the sentences after the table of contents where we
        # find the title in the text. Note, once we find the title, we take all the sentences after that set by the
        # increment variable.
        for sentence in sentenceList[wordLocation+1:]:  # looks at each sentence after the table of contents.
            cleanSentence = clean(sentence)  # Use the clean function to avoid spaces and new lines affecting code.
            if title.lower() in cleanSentence.lower():  # checks if the cleaned title is in the cleaned sentence.
                loc = sentenceList.index(sentence)  # finds the index of the start of the section in the sentence list
                potential.append(sentenceList[loc:loc+increment])  # adds the increment in a list of sections called
                # potential

        #  below are all the potential formats of page numbers.
        str1 = str(page1) + " \n"
        str2 = str(page2) + " \n"
        str3 = str(page1) + "\n"
        str4 = str(page2) + "\n"
        str5 = str(page1) + "-\n"
        str6 = str(page2) + "-\n"
        str7 = str(page1) + " -\n"
        str8 = str(page2) + " -\n"

        # this is a buffer that avoids the below check since it means that there is only one section after the table of
        # contents with the word "Background in it"
        if len(potential) == 1:
            final = potential[0]

        # the below 50 lines may look intimidating but essentially it first takes the list of sentences and makes them
        # into strings in order to simplify the parsing. Then since the starting page number comes before the ending
        # page number, I look for the starting number in the string first and then ensure that the ending page number is
        # in the section.
        for section in potential:
            empty = ''
            for sentence in section:
                empty += sentence  # adds the sentence to the string.
            if str1 in empty:  # check
                if str2 in empty:
                    final = section
                    break
                elif str4 in empty:
                    final = section
                    break
                elif str6 in empty:
                    final = section
                    break
                elif str8 in empty:
                    final = section
                    break
            elif str3 in empty:
                if str2 in empty:
                    final = section
                    break
                elif str4 in empty:
                    final = section
                    break
                elif str6 in empty:
                    final = section
                    break
                elif str8 in empty:
                    final = section
                    break
            elif str5 in empty:
                if str2 in empty:
                    final = section
                    break
                elif str4 in empty:
                    final = section
                    break
                elif str6 in empty:
                    final = section
                    break
                elif str8 in empty:
                    final = section
                    break
            elif str7 in empty:
                if str2 in empty:
                    final = section
                    break
                elif str4 in empty:
                    final = section
                    break
                elif str6 in empty:
                    final = section
                    break
                elif str8 in empty:
                    final = section
                    break

        finalList = []
        for sentence in final:  # this loop checks the position of page number of the next section inorder to end the string
            if str2 in sentence:
                index = final.index(sentence)  # looks for the ending page number in the sentence.
                finalList = final[:index]  # takes everything from the start to the ending page number's occurrence
            elif str4 in sentence:
                index = final.index(sentence)
                finalList = final[:index]
            elif str6 in sentence:
                index = final.index(sentence)
                finalList = final[:index]
            elif str8 in sentence:
                index = final.index(sentence)
                finalList = final[:index]

        stringFinal = ''
        for sent in finalList:  # this loop is converting the list of sentences into a string in order to put into excel
            stringFinal = stringFinal + sent
        stringFinal = clean(stringFinal)  # removes all extra spaces and newlines
        filenames.append(stringFinal)
        if len(finalList) == 0:
            excel.write_data(excel_path, "Sheet1", row, 10, "Error")  # adds error into the excel
            print("Error")
        else:
            excel.write_data(excel_path, "Sheet1", row, 10, stringFinal)  # add the string if its correct

d = {}  # this is a dictionary to store all the useful words.
for file in filenames:  # checks each file in the file names variable.
    data = clean(file)
    data = data.split()
    value = 0
    for i in data:
        if i.isalpha() == True:
            if i not in stop_words:
                i = i.lower()
                if i in d:
                    d[i] += 1
                else:
                    d[i] = 1

        # updates the values of the main dictionary.
for w in sorted(d, key=d.get, reverse=True):
    print(w, d[w])  # this just prints the word and its number of occurrences

        # for word in bag:  # this is counting the number of occurrances of a word.
        #     if word in d.keys():
        #         print(word, " has ", d.get(word), " occurrences")