def fileread(file):
    txt = open(file)
    print(txt.read())

def index(filepath, keywords):
    with open(filepath) as f:
        for lineno, line in enumerate(f, start=1):
            matches = [k for k in keywords if k in line]
            if matches:
                result = "{:<15} {}".format(','.join(matches), lineno)
                print(result)


fileRead('cores.txt')
fileRead('carta.txt')
index('cores.txt', ["amarelo"])
index('carta.txt', ["azul"])
index('carta.txt', ["verde"])