game.autorun

tile.style 2
pathway.style 2

# clear everything
Object.clear
Pathway.clear
Message.clear
Block.clear
event.clear
pathway.block 1 

a = 0 
for a = 0 to 31
    object.destroy 28, a
    object.destroy 8, a	
    object.destroy 6, a
next 

for a = 1 to 3
    event.destroy 8, a
next 

event.create "plarn08", 50, 6, 11
object.create 6, 0, 6, 11

hero.create "Harry Potter", 8,2
Hero.steps 2,2,0,2,0,0,0,0,0,1,1,1,1   # auxiliary instruction tool For demonstrating the wizard ' s walking route
hero.say "I need new python command.", 1
Message.show "Level 8 \n Mission: Destroy the wall and get out. \n Hint: \n message.close \n  a = 0 \n  for a = 4 to 9  \n       Pathway.destroy 6, a \n       Pathway.destroy 7, a \n       Pathway.destroy 8, a \n       Pathway.destroy 9, a \n  next ", 6
Message.show " The wizard in the Dark Forest has broken the door and trapped you. But you are the best student in this class! You can destroy the wall and get out!  ", 7

Object.create 25, 0, 1, 0
Object.create 25, 1, 0, 1
Object.create 19, 0, 6, 9
Object.create 24, 0, 4, 1
Object.create 25, 2, 11, 6

Object.create 1, 0, 0, 3
#Object.create 9, 0, 9, 4
Object.create 25, 0, 7, 1
Object.create 25, 1, 9, 2
Object.create 25, 2, 11, 1
Object.create 25, 3, 12, 1
Object.create 29, 0, 13, 11 # three color flower
Object.create 18, 4, 13, 10  
Object.create 18, 5, 12, 10  
Object.create 28, 0, 2, 1
Object.create 24, 1, 1, 1

Pathway.draw 0, 0, 16, 14
Hero.create "Mike", 8, 3 
Event.create 7, 8, 3

Hero.steps 2, 2, 2, 0, 0

a = 0
b = 5
c = 0

For a = 8 To 13
        Pathway.destroy 1, a
Next  

For a = 7 To 13
        Pathway.destroy 5, a
Next  

For a = 5 To 5
        Pathway.destroy 10, a
Next

For a = 9 To 10
        Pathway.destroy 10, a
Next  
For a = 7 To 9
        Pathway.destroy 14, a
Next  

For a = 1 To 10
        Pathway.destroy a, 11
Next  

For a = 5 To 10
        Pathway.destroy a, 7
Next  

For a = 10 To 14
        Pathway.destroy a, 9
Next

For a = 14 To 15
        Pathway.destroy a, 7
Next  

For a = 6 To 14
    For b = 0 To 4
        Pathway.destroy a, b
    Next  
Next








