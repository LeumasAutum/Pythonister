def substitution_cipher(text, key):
encrypted_text = ""
for char in text:
if char.isalpha():
char_code = ord(char)
char_code += key
if char.isupper():
if char_code > ord('Z'):
char_code -= 26
encrypted_text += chr(char_code)
else:
if char_code > ord('z'):
char_code -= 26
encrypted_text += chr(char_code)
else:
encrypted_text += char
return encrypted_text

text = input("Enter the text you want to encrypt: ")
key = 7

print("You entered:", text)
print("Your encrypted text is:", substitution_cipher(text, key))
