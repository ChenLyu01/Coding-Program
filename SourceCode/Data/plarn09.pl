game.autorun

Tile.style 2
Pathway.style 2

# clear everything
Object.clear
Pathway.clear
Message.clear
Block.clear
event.clear
pathway.block 1 

Hero.create "Mike", 0, 10
Event.create 7,0 , 10


Hero.steps 2, 2, 3, 3, 2, 2, 2, 2, 3, 3, 2, 2, 2, 2, 3, 3, 2, 2, 2, 2, 3, 3, 2, 2
hero.say "I see Mandragora!", 1
Message.show "Level 10 \n Mission: Go to the Mandragora Field and \n Open the door with limited magic energy. \n Hint: \n message.close \n  a = 0 \n  for a = 0 to 3  \n      hero.moveright 2 \n      hero.moveup 2 \n      hero.moveright 2 \n  next ", 6
Message.show "There is a pattern! I need to save my energy by using loops! Ask your teacher what is a loop? Hint: \n  a = 0 \n  for a = 0 to 3  \n      hero.moveright 2 \n      hero.moveup 2 \n      hero.moveright 2 \n  next ", 7

event.create "plarn10", 50, 15, 2
object.create 6, 0, 15, 2


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
Object.create 24, 0, 11, 2
Object.create 24, 1, 10, 1
Object.create 24, 2, 10, 2
Object.create 27, 0, 9, 1
Object.create 24, 3, 2, 6




