import re


a = "False Call Device (%)"


print(re.sub(r" \(\W+\)", "", a))