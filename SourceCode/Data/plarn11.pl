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

Hero.create "Mike", 15, 2
Hero.steps 1, 1, 1, 0, 0, 1, 1, 1, 1, 0, 0, 1, 1, 1, 1, 0, 0, 1, 1, 1, 1, 0, 0, 1
Event.create 7,15, 2

hero.say "What potion is Professor V making?", 1
Message.show "Level 12 \n Mission: Head back quickly with loops magic spells\n Hint: \n Hero.moveleft \n Hero.movedown .)", 6
Message.show "You finally get all the ingredients needed for the potion. But you are almost out of magic power! Get back as fast as you can with fewer lines of codes,)", 7

Event.create "plarn12", 50, 1, 10
Object.create 6, 0, 1, 10

Pathway.draw 0, 0, 16, 14  

Object.create 30, 0, 15, 6
Object.create 30, 1, 13, 6
Object.create 30, 2, 11, 6
Object.create 25, 0, 0, 1
Object.create 25, 2, 2, 0
Object.create 25, 3, 3, 0


m = - 3
n = 10
For q = 0 To 4
       d = n
For a = 0 To 4
       c = a + m
       Pathway.destroy c, d
Next

For b = 0 To 2
      d = n - b
      Pathway.destroy c, d
Next  
       m = m + 4
       n = n - 2
Next  

For a = 0 To 5
    For b = 0 To 3
          Pathway.destroy a, b
      Next  
  Next

For a = 11 To 15
      For b = 6 To 8
           Pathway.destroy a, b
       Next  
   Next  

   For b = 3 To 5
       Pathway.destroy 15, b
   Next

Object.create 9, 0, 0, 3
Object.create 9, 1, 3, 3
Object.create 9, 2, 11, 8
Object.create 9, 3, 14, 8
Object.create 9, 4, 11, 6
Object.create 21, 0, 6, 9
Object.create 16, 0, 3, 11
Object.create 16, 1, 12, 10
Object.create 17, 0, 3, 6
Object.create 17, 1, 2, 5
Object.create 24, 0, 11, 2
Object.create 24, 2, 10, 2
Object.create 24, 1, 9, 2
Object.create 24, 3, 9,1


