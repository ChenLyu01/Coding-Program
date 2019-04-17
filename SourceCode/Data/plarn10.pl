Game.autorun
Tile.style 2
Pathway.style 2

# clear everything
Object.clear
Pathway.clear
Message.clear
Block.clear
event.clear
pathway.block 1 

Hero.create "Mike", 8, 4
Event.create 7, 8, 4
Hero.steps 0, 0, 0, 0, 0, 1, 1, 1, 3, 3, 2, 2, 2, 2, 2, 0, 0, 1, 1,1
hero.say "I am the best!", 1
Message.show "Level 11 \n Mission: Open the door and create more Mandragora! \n Please type: \n object.create \n hero.facedown \n hero.open \n hero.movedown", 6
Message.show "Oh my god! The Chomper has eaten so many Mandragora! Professor has just told me to grow more Mandragora!  ", 7

event.create "plarn11", 50, 3, 9
object.create 6, 0, 3, 9

Pathway.draw 0, 0, 16, 14
 
Object.create 9, 0, 7, 5
Object.create 24, 0, 1, 1
Object.create 16, 0, 1, 8
Object.create 16, 1, 1, 11

Object.create 30, 1, 5, 7
Object.create 30, 2, 10, 6
Object.create 15, 1, 5, 5

Object.create 15, 2, 12, 6
Object.create 24, 1, 13, 1
Object.create 24, 2, 13, 11
Object.create 24, 3, 11, 11
a = 0
b = 5
c = 0
For c = 0 To 2
    Pathway.destroy c, b
Next
b = 3
For c = 2 To 6
    Pathway.destroy c, b
Next

For c = 10 To 14
    Pathway.destroy c, b
Next

b = 5
For c = 14 To 15
      Pathway.destroy c, b
Next
b = 1
For c = 6 To 10
      Pathway.destroy c, b
Next

For c = 1 To 4
    Pathway.destroy 8, c
Next  

Pathway.destroy 2, 4
Pathway.destroy 6, 2
Pathway.destroy 10, 2
Pathway.destroy 14, 4

For a = 4 To 12
     For b = 5 To 9  
     Pathway.destroy a, b
    Next  
Next  

For a = 3 To 13
     For b = 6 To 9  
     Pathway.destroy a, b
    Next  
Next
