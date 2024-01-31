import re
import xlsxwriter

wordCounter = {}
personMessageCounter = {}
personWordCounter = {}

lines = []

people = {
    # here you can put the people in the chat that are not on your phone book or
    # weren't on you phone book at the time of the chat.
    # you can also put synonyms for people
    # example: "<phone number>":"nickname"
}


def resolve_person(person):
    if person.find(' ') == 0:
        person = person[1:]
    if person[-1] == ' ':
        person = person[:-2]
    person = re.sub(r'[^\d\s\w+\-:,\.א-ת]', '', person)

    if person in people:
        return people[person]
    else:
        print('"'+person+'"')
        return person


with open(f'chat.txt', 'r', encoding='utf-8') as f:
    for line in f:
        timeEndIndex = line.find('-')
        if timeEndIndex == -1 or re.match('^(\d{1,2}\.){2}\d{4}, \d{1,2}:\d{1,2} $', line[:timeEndIndex]) is None:
            message = re.sub(r'[,\-\.@\d\n]', ' ', line)
            message = re.sub(r'[^\w| ]', '', message)
            lines[-1]['message'] += message
        else:
            time = line[:timeEndIndex]
            content = line[timeEndIndex + 1:]
            colonIndex = content.find(':')

            if colonIndex == -1:
                # system message
                pass
            else:
                # regular message
                sender = resolve_person(content[:colonIndex])
                message = content[colonIndex + 1:]
                message = message.replace('<המדיה לא נכללה>', ' ')
                message = message.replace('הודעה זו נמחקה', ' ')
                message = re.sub(r'[,\-\.@\d\n]', ' ', message)
                message = re.sub(r'[^\w| ]', '', message)

                lines.append({'message': message, 'sender': sender, 'time': time})

for line in lines:
    words = [word for word in line['message'].split(' ') if word != '']

    if line['sender'] in personMessageCounter:
        personMessageCounter[line['sender']] += 1
    else:
        personMessageCounter[line['sender']] = 1

    if line['sender'] in personWordCounter:
        personWordCounter[line['sender']] += len(words)
    else:
        personWordCounter[line['sender']] = len(words)

    for word in words:
        if word in wordCounter:
            wordCounter[word] += 1
        else:
            wordCounter[word] = 1


def print_dict(dict):
    keysList = list(dict.keys())
    keysList.sort(key=lambda word: dict[word], reverse=True)
    for word in keysList:
        print(word + ': ' + str(dict[word]))

def write_to_xls(worksheet, dict, key_column, value_column):
    keysList = list(dict.keys())
    keysList.sort(key=lambda key: dict[key], reverse=True)

    row = 1

    for key in keysList:
        worksheet.write(key_column + str(row), key)
        worksheet.write(value_column + str(row), dict[key])
        row += 1



print_dict(personWordCounter)

workbook = xlsxwriter.Workbook('chat.xlsx')
worksheet = workbook.add_worksheet('analysis')

write_to_xls(worksheet, personWordCounter, 'A', 'B')
write_to_xls(worksheet, personMessageCounter, 'C', 'D')
write_to_xls(worksheet, wordCounter, 'E', 'F')

workbook.close()
