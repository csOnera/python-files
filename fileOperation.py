from datetime import datetime
import time

thisYear = datetime.now().year #integer


def create_file(filename=r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\year.txt"):
    try:
        with open(filename, 'w') as f:
            f.write(str(thisYear))
        print("File " + filename + " created successfully with year.")
    except IOError:
        print("Error: could not create file " + filename)
 
def read_file(filename= r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\DATABASE_TRIAL\python files\year.txt"):
    try:
        with open(filename, 'r') as f:
            year = f.read()
            return year
    except IOError:
        print("Error: could not read file " + filename)
 
def append_file(filename, text):
    try:
        with open(filename, 'a') as f:
            f.write(text)
        print("Text appended to file " + filename + " successfully.")
    except IOError:
        print("Error: could not append to file " + filename)


if __name__ == '__main__':
    try:
        year = read_file()
    except:
        create_file("year.txt")
        print(f"Reminder: the newly created file is with year {thisYear}, if now is {thisYear} but before 1st of April, please edit the year manually back to {thisYear - 1}")

    print(f"Current year in record is {year if year != '' else 'not found'}")

    while True:
        answer = input(f"input 'change year' to turn the year to {thisYear}:  ")

        if answer == 'change year':
            create_file()
            # print(f"successfully changed year to {thisYear}")
            time.sleep(10)
            exit()
        else:
            print("please input valid answer")

