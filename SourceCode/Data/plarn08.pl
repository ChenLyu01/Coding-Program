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

Hero.create "Harry Potter", 11, 3  
Event.create 7, 11, 3

Hero.steps 0, 0, 1, 1, 1, 1, 1, 0, 0, 0, 0, 2, 2, 2, 2, 2, 3, 3, 2, 2, 2
hero.say "I have to cast spells faster. ", 1
Message.show "Level 9 \n Mission: Go to the Mandragora Field. \n Hint: \n message.close \n  a = 0 \n  for a = 4 to 9  \n       Pathway.destroy 6, a \n       Pathway.destroy 7, a \n       Pathway.destroy 8, a \n       Pathway.destroy 9, a \n  next ", 6
Message.show "The second ingredient needed is Mandragora! The magic stone is guiding the way to find it. Persist! ", 7

event.create "plarn09", 50, 15, 6
# object.create 6, 0, 6, 11

Pathway.draw 0, 0, 2, 14  
Pathway.draw 3, 0, 8, 4
Pathway.draw 3, 4, 3, 4
Pathway.draw 3, 9, 3, 5
Pathway.draw 7, 5, 4, 3
Pathway.draw 7, 9, 9, 5
Pathway.draw 12, 7, 4, 2
Pathway.draw 12, 0, 3, 6
Pathway.draw 15, 0, 1, 4
Pathway.draw 15, 0, 1, 4
Pathway.draw 11, 5, 1, 1

Object.clear
object.create 6, 4, 15, 6
Object.create 1, 3, 0, 0 # gate
Object.create 19, 3, 7, 1 # tree base  
Object.create 21, 3, 7, 9 # tree
Object.create 9, 0, 10, 0 #
Object.create 28, 0, 11, 12
Object.create 28, 1, 13, 12
Object.create 25, 0, 13, 3
Object.create 22, 0, 12, 3
Object.create 22, 1, 13, 2
Object.create 17, 0, 4, 10
Object.create 17, 1, 3, 10
Object.create 22, 2, 8, 6

#answer
#Hero.movedown  
#Hero.moveleft 5
#Hero.movedown 4
#Hero.moveright 5
#Hero.moveup 2
#Hero.moveright 4




