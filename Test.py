import re

# sample sentence
sentence = "The price of the item is $25.50 and the quantity"

# regular expression pattern to match numbers
pattern = r'\d+\.?\d*'

# extract numbers from the sentence
numbers = re.findall(pattern, sentence)
numbers = numbers[0]

# print the extracted numbers
print(numbers)
