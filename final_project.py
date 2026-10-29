#   Password strength checker

import re

# password strength check conditions :
#  min 8  charas, digit, uppercase,lowercase, & special chars

def check_password_strength(password):
          if len(password)<8  :   # length of password
            return "week: password must be at least 8 chars"
          
          if not any(char.isdigit() for char in password) :
                return  "week : password must contain a digit"
          
          if not any(char.isupper() for char in password) :
                return "week: password must be at least 8 chars"
          
          if not any(char.islower for char in password):
                return "week : password must contain a digit"
          
          if not re.search(r'[!@#$%^&*(){}<>.?]',password) :
                return "Medium : password must contain a special charc"
          
          return "Strong : Your password is secured"

def  password_checker():
      print("Welcome to the password strength check")

      while True   :
            password=input("enter your password (or type 'exist' to quite):-")

            if password.lower() == exit:
                  print ("Thanks you using this tool")
                  break       
            
            result= check_password_strength(password)
            print(result)

#  Run the password_checker tool
if __name__=="__main__":
      password_checker()            
 

