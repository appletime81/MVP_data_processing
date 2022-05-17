import collections


def main():
    a = "hackthegame"
    counter = collections.Counter(a)
    print(counter)


    if 1 not in counter.values():
        return -1
    else:
        key = sorted([k for k, v in counter.items() if v == 1])[0]

        print(a.index(key) + 1)

main()