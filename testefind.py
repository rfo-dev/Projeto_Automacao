import fitz

# defining string 
str1 = "zzzAzzzAbazzzA"
  
# defining substring
substr = "Aba"
  
# printing original string 
print("The original string is : " + str1)
  
# printing substring 
print("The substring to find : " + substr)
  
# using list comprehension + startswith()
# All occurrences of substring in string 
res = [i for i in range(len(str1)) if str1.startswith(substr, i)]
  
# printing result 
print("The start indices of the substrings are : " + str(res))