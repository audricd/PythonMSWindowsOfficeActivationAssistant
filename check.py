keytoinstall = input("write licence:")


def lencheck():
    if len(keytoinstall) != 23:
        return False
    else:
        return True


def sepcheck():
    if keytoinstall[5] == "-" and keytoinstall[11] == "-" and keytoinstall[17] == '-':
        return True
    else:
        return False


def charcheck():
    if str.isalnum(keytoinstall[0:4]) and str.isalnum(keytoinstall[6:10]) and str.isalnum(keytoinstall[12:17]):
            return True
    else:
        return False

if lencheck() == True and sepcheck() == True and charcheck() == True:
    validlicence = True
else:
    validlicence = False

print(validlicence)
