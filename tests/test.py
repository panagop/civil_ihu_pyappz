# Write your frequency_dictionary function here:
def frequency_dictionary(words):
  d = {}
  for word in words:
    print(f'word = {word}')
    if word in d.keys():
      print('True: word in d')
      d[word] +=1
      
    else:
      print('False: word not in d')
      d[word] = 1
    print(d)
  return d


# Uncomment these function calls to test your  function:
print(frequency_dictionary(["apple", "apple", "cat", 1]))
# should print {"apple":2, "cat":1, 1:1}
print(frequency_dictionary([0,0,0,0,0]))
# should print {0:5}