import re
import nlp
import excel
from nlp import sent_tokenize
from nlp.corpus import stopwords

excel_path="/Users/sameerhimati/Desktop/ISBStuff/All.xlsx"
num_rows=excel.get_rows(excel_path,"Sheet1")
num_cols=excel.get_columns(excel_path,"Sheet1")
nlp.download("punkt")
nlp.download("stopwords")
stop_words = set(stopwords.words('english'))


def makeSections(filename, phrase1, phrase2):
    """
    :param filename:  TextFile
    :param phrase1: String
    :param phrase2: String
    :return: Tuple of lists, first index of the tuple is a list of pages. The second tuple is a list of indices of the pages.

    This function takes the 2 phrases provided and looks for them in a file. Then it ensures that it only adds
    """
    file = open(filename, 'r')  # opens the file in read mode
    data = file.read().split("<PAGE>")  # Splits the file into a list of pages and stores it in a variable called data.

    sections = []  # this is a variable used to store sections.

    index1 = []  # this variable stores the all the pages with phrase1 in it
    for page in data:  # iterates over each page in the file
        page = page.lower()
        if page.find(phrase1.lower()) != -1:  # checks if phrase1 was found in that page
            index1.append(page)  # adds the page containing phrase1 to index1
    index2 = []  # this variable stores all the pages with phrase2
    for page in data:  # iterates over each page
        page = page.lower()
        if page.find(phrase2.lower()) != -1:  # checks if phrase2 was found in that page
            index2.append(page)  # adds the page containing phrase2 to index2
    indices1 = []  # this variable stores the indices of all the pages with phrase1
    for page in data:
        for i in range(len(index1)):  # iterates over the list containing the indices of all the pages with phrase1.
            if page.lower().find(index1[i].lower()) != -1:  # checks if the page in found in data, in case, extra check
                indices1.append(data.index(page))  # if phrase1 was found, adds the index of that page with phrase1
                # (Note not the page itself, just the index)
    indices2 = []  # this variable stores the indices of all the pages with phrase2
    for page in data:
        for i in range(len(index2)):  # iterates over the list containing the indices of all the pages with phrase2.
            if page.lower().find(index2[i].lower()) != -1:  # checks if the page in found in data, in case, extra check
                indices2.append(data.index(page))  # if phrase2 was found, adds the index of that page with phrase2
                # (Note not the page itself, just the index)
    indices = []
    for index in indices1:
        print(index)
        diff = (min_diff(index, indices2) - index) + 1  # this calculation uses the below helper function
        # it helps find the size of how long the section needs to be.
        pages = []  # this variable will store all the pages needed.
        if diff > 10:  # a filter to remove unwanted pages as Background of the merger will almost never be over 5 pages
            print("section over 10 pages found.")
            pass
        for j in range(diff):  # iterates over the difference so if the difference is 5 starts at 0 and goes up to 5.
            pages.append(data[index + j])  # adds all the pages from the first occurrence of phrase1 till phrase2.
            indices.append(index + j)  # also adds its indices for future functions.
        sections.append(pages)  # note sections is a list of list of strings. As its a list of sections.

    return (sections, indices)  # returns a tuple containing a list of pages and its indices.


def min_diff(num, l):  # used to find where the num variable is in the list l.
    print(l)
    l.append(num)  # first adds the number mun to the list l.
    l.sort()  # then sorts the list l using python's default sort function.
    ind = l.index(num)  # the new ind variable stores the index of where the number num is in the list l.
    return l[ind + 1]  # returns the next index (page number) after the given one.


def findWords(filename, phrase1, phrase2):
    """
    :param filename: TextFile
    :param phrase1: String
    :param phrase2: String
    :return: Dictionary of words

    The purpose of this function is to return a dictionary of words and the number of times they occur between the
    sections from phrase1 to phrase2. Note this function uses the makeSections function we wrote previously as a helper.
    """
    file = open(filename, "r")  # opens the file in read mode
    data = file.read().split("<PAGE>")  # once again splits it into a list of strings by page number.
    dic = {}  # initializing a new dictionary variable, dic.
    (section, index) = makeSections(filename, phrase1, phrase2)  # calls upon the make section function and stores it.
    for i in index:  # using the indices of the pages, iterates over them.
        page = data[i]  # associates the page variable to a particular page in the document.
        paragraphList = page.split("\n\n")  # this splits the page into a list of paragraphs.
        for paragraph in paragraphList:  # for each paragraph in a list of paragraphs.
            lineList = paragraph.split()  # splits them into a list of lines.
            for line in lineList:  # checks each line.
                clean = re.sub('[^a-zA-Z]+', '', line.lower())  # this uses the regular expression method of python.
                # It only includes all Alphabetic characters.
                if clean not in stop_words:  # stop words is a list of generic words that we see often.
                    # the list can be found by typing print(stop_words) in console. This ensures these are filtered.
                    if clean.lower() not in dic:  # checks if its the first occurrence of the word.
                        dic[clean.lower()] = 0  # if it is then its added to the dictionary.
                    dic[clean.lower()] += 1  # and then the word's(key) value is incremented by one.
    return dic  # returns the dictionary.


def makeParagraph(Pagelist):
    """
    :param Pagelist: List of sections. (normally len(PageList) = 1)
    :return: divides the pages into paragraphs

    Takes the input list and divides it into paragraphs. Returns a list of paragraphs.
    """
    paras = []  # a variable to store paragraphs.
    for section in Pagelist:  # iterates over each section
        for page in section:  # iterates over each page in section as its already divided by pages.
            paraList = page.split("\n\n")  # makes a list of paragraphs
            for para in paraList:  # iterates over each paragraph in the list of paragraphs in that page
                paras.append(para)  # adds each paragraph into the return list.  # removes newline characters.
    return paras


def selectParagraphs(sentenceList, bag):
    """
    :param sentenceList: List of  String, A list of sentences.
    :param bag: List of strings, The bag of words looking to filter (All lower case and spaces before and after)
    :return: List of useful sentences.

    This function takes a list of sentences as input along with the bag of words we are looking for and filters
    them
    """
    useful = []  # A list to store important sentences.
    for sentence in sentenceList:  # iterates over each sentence
        sentence = sentence.lower()  # adjusts for case.
        for word in bag:  # iterates over the bag of words.
            if word in sentence:  # check to see if that word occurs in the sentence
                if sentence not in useful:  # in order to avoid repetition of sentences.
                    # Since a sentence can contain 2 words from the bag.
                    useful.append(sentence)  # adds only the sentences with a word from the bag.
    return useful


def sameSentence(para, word1, word2):
    """
    :param para: String, Takes a paragraph as input.
    :param word1: String, Acquirer name with spaces before and after
    :param word2: String, Target Name with spaces before and after
    :return: Integer, -1 or 1.

    This function takes a paragraphs and 2 words,
    returns whether word1 and word2 are in the same sentence in the paragraph.
    """
    i = -1  # indicator.
    word1 = word1.lower()  # make lowercase since the in function in python is case sensitive.
    word2 = word2.lower()
    SentList = sent_tokenize(para)  # Using NLTK sent_tokenize function, divides the paragraph into a list of sentences.
    for sentence in SentList:  # iterating over the sentences.
        if ((word1 in sentence.lower()) & (word2 in sentence.lower())):  # checks if both words are in the sentence.
            i = 1  # if they are then sets i = 1.
    return i  # returns whether or not there is a sentence in the paragraph that has both words.


def giveSentence(para, word1, word2):
    """
    :param para: String, Paragraph as input.
    :param word1: String, Acquirer name with spaces before and after
    :param word2: String, Target Name with spaces before and after
    :return: list of sentences with both words in them

    Takes the paragraph as input and returns a list of sentences from that paragraph that contain the both words.
    """
    sen = []  # variable to store sentences.
    word1 = word1.lower()  # make lowercase since the in function in python is case sensitive.
    word2 = word2.lower()
    SentList = sent_tokenize(para)  # Using NLTK sent_tokenize function, divides the paragraph into a list of sentences.
    for sentence in SentList:  # iterating over the sentences.
        if ((word1 in sentence.lower()) & (word2 in sentence.lower())):  # checks if both words are in the sentence.
            sen.append(sentence)  # adds the sentence to the list of important sentences.
    return sen  # returns the list of sentences.


def RunCode(file, acquirer, target):
    phrase1 = "Background of the "
    phrase2 = "Reasons for the "
    bag = [" met ", " discussions ", " agreed ", " meeting ", " proposed "]
    sectionlist = makeSections(file, phrase1, phrase2)
    filtered = []
    sen = []  # this variable is important as it stores all the important sentences.
    paragraphs = makeParagraph(sectionlist[0])
    for para in paragraphs:  # goes over each paragraph.
        sentencelist = giveSentence(para, acquirer, target)  # stores all sentences that contain the acquirer and target
        if len(sentencelist) > 0:  # ensures that the list isn't empty.
            for sentence in sentencelist:  # iterates over the list of useful sentences
                sen.append(sentence)  # and adds them to the placeholder variable sen.
    filteredsentences = selectParagraphs(sen, bag)
    for sentence in filteredsentences:
        filtered.append(sentence)
    return filtered


def filterWords(phrase):
    stopwords = ['co', 'inc', 'corp', 'ltd', 'Products']
    phraselist = phrase.split()

    resultwords = [word for word in phraselist if word.lower() not in stopwords]
    result = ' '.join(resultwords)
    return result


for row in range(2, 237):
    useful = []
    deal = excel.read_data(excel_path, "Sheet1", row, 1)
    acquirer = excel.read_data(excel_path, "Sheet1", row, 9)
    target = excel.read_data(excel_path, "Sheet1", row, 10)
    txt = str(deal)
    txt = txt + ".txt"
    a = filterWords(acquirer)
    t = filterWords(target)
    filtered = RunCode(txt, a, t)
    useful.append(filtered)
    print(filtered)
    print(len(filtered))


# filenames = ["14AExxonMobil.txt"]  # used to find the frequency of words.
# filename = "14AExxonMobil.txt"  # input text file.
# phrase1 = "Background of the Merger\\\n\\\n"  # section to start.
# phrase2 = "Reasons for the Merger\\\n\\\n"  # section end.
# # filename = input("Give me a filename with a .txt extension included: ")
# # phrase1 = input("I also need a phrase you are looking for: ")
# # phrase2 = input("AND another lol: ")
#
# acquirer = " Exxon "  # acquirer name
# target = " Mobil "  # target name
# bag = [" met ", " discussions ", " agreed ", " meeting ", " proposed "]  # bag of words to filter.
#
# sectionList = makeSections(filename, phrase1, phrase2)
# # don't need the below either
# #
# #
# # ---------------------------------------------------------------
# for section in sectionList[0]:
#     print("Number of Pages in Section is ", len(section))
#     for page in section:
#         print(page)
#         print("+++++")
# # ---------------------------------------------------------------
# # ---------------------------------------------------------------
# # makes paragraphs below and only returns the useful paragraphs.
# #
# #
# # Basically makes the pages into sections and divides it into useful paragraphs then  filters the useful paragraphs by
# # words in the bag.
# # don't really need this
# useful = []
# paragraphs = makeParagraph(sectionList[0])
# for para in paragraphs:
#     if sameSentence(para, acquirer, target) > 0:
#         useful.append(para)
# print(len(useful))
#
# for a in useful:
#     print(a)
#     print("******************")
#
# # prints the useful paragraphs.
# filteredParagraph = selectParagraphs(useful, bag)
# print(len(filteredParagraph))
# for para in filteredParagraph:
#     print(para)
#     print("~~~~~~~~~~~~~~~~~~~~~~")
# #
# #
# #
# # ---------------------------------------------------------------
# #
# #
# #
# # makes sentences below
# sen = []  # this variable is important as it stores all the important sentences.
# paragraphs = makeParagraph(sectionList[0])
# for para in paragraphs:  # goes over each paragraph.
#     sentenceList = giveSentence(para, acquirer, target)  # stores all sentences that contain the acquirer and target
#     if len(sentenceList) > 0:  # ensures that the list isn't empty.
#         for sentence in sentenceList:  # iterates over the list of useful sentences
#             sen.append(sentence)  # and adds them to the placeholder variable sen.
#
# filteredSentences = selectParagraphs(sen, bag)  # this filters the sentences with the bag of words and returns only the
# # ones containing words from the bag.
#
# print("There are ", len(filteredSentences), " filtered sentences.")
# for sentence in filteredSentences:
#     print(sentence)
#     print("...........................")
#
#
#
# d = {}  # this is a dictionary to store all the useful words.
# for file in filenames:  # checks each file in the file names variable.
#     temporary_Dictionary = {}  # initiates a new dictionary to act as a placeholder dictionary.
#     temporary_Dictionary = findWords(file, phrase1, phrase2)  # adds the dictionary of words to the placeholder Dic.
#     d.update(temporary_Dictionary)  # updates the values of the main dictionary.
#
# for w in sorted(d, key=d.get, reverse=True):
#     print(w, d[w])  # this just prints the word and its number of occurrences
#
# for word in bag:  # this is counting the number of occurrances of a word.
#     if word in d.keys():
#         print(word, " has ", d.get(word), " occurrences")

