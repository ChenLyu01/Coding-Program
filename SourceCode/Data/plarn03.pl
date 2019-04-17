Game.autorun  

Tile.style 2
Pathway.style 2

# clear everything
Object.clear
Pathway.clear
Message.clear
Block.clear
Event.clear
Pathway.block 1  


Hero.create "Harry Potter", 11, 6  
Event.create 7, 11, 6

  
Hero.steps 0, 0, 0, 1, 1, 1, 1, 1, 3, 3, 3, 3, 2, 2, 2, 2, 2, 3, 3
Pathway.draw 0, 0, 2, 14  

Pathway.draw 3, 0, 8, 4

Pathway.draw 3, 4, 3, 4
Pathway.draw 3, 9, 3, 5
Pathway.draw 7, 5, 4, 3
Pathway.draw 7, 9, 9, 5
Pathway.draw 12, 7, 4, 2
Pathway.draw 12, 0, 3, 6
Pathway.draw 15, 0, 1, 4
Pathway.draw 11, 5, 1, 1


Hero.say "I am the best student", 1
Message.show "Task: \n Please type: \n hero.movedown \n hero.moveleft \n hero.moveup \n hero.moveright", 6
Message.show "You know, there's a witch in the magic school. She always rewards the best students and helps them realize their dreams of life. ", 7


Object.clear
Object.create 1, 3, 0, 0 # gate
Object.create 6, 3, 11, 1 # magic stone
Object.create 19, 1, 6, 1 # tree base  
Object.create 21, 3, 7, 9 # tree
Object.create 22, 1, 4, 12 # yellow flower
Object.create 22, 2, 8, 6 # yellow flower
Object.create 25, 0, 13, 8 # grass
Object.create 9, 0, 10, 0 #
Object.create 9, 2, 7, 0 #
Object.create 9, 3, 13, 0 #

Event.create "plarn04", 50, 11, 1
