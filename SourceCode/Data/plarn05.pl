game.autorun

Tile.style 2
Pathway.style 2

# clear everything
Object.clear
Pathway.clear
Message.clear
Block.clear
Event.clear
Pathway.block 1  

Hero.create "Harry Potter", 2, 3
Event.create 7, 2,3
Hero.steps 0, 0, 0, 0, 0, 0, 2, 2, 2, 2, 3, 3, 3, 3, 2, 2, 2, 2, 2, 3, 3


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


Hero.say  "Hurry up! I need to move fast!", 1
Message.show "Level 6 \n Task: Go back to the Knotgrass Field. Hurry!\n Hint: \n hero.moveright () \n hero.movedown () \n hero.moveup ()", 6
Message.show "I can save my magical energy and cast spells faster by using fewer lines of codes.", 7

Object.clear
Object.create 1, 3, 0, 0 # gate
Object.create 6, 4, 11, 1 # magic stone
Object.create 19, 1, 6, 1 # tree base  
Object.create 21, 3, 7, 9 # tree
Object.create 22, 1, 4, 12 # yellow flower
Object.create 22, 2, 8, 6 # yellow flower
Object.create 25, 0, 13, 8 # grass
Object.create 9, 0, 10, 0 #
Object.create 9, 2, 7, 0 #
Object.create 9, 3, 13, 0 #

Event.create "plarn06", 50, 11, 1

#Answer
#Hero.movedown 5
#Hero.moveright 4
#Hero.moveup 4
#Hero.moveright 5
#hero.moveup 3
