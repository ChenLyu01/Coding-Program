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



Hero.create "Harry Potter", 11, 1
Event.create 7, 11,1

hero.steps 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 0, 0, 1, 1, 1, 1, 3, 3, 3
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

Hero.say  "Oh, I forget the key...", 1
Message.show "Level 5 \n Mission: Hurry back to the door to get the key.\n Hint: Get back home with only 5 lines of codes \n hero.movedown \n hero.moveleft \n hero.moveup", 6
Message.show "No one can succeed easily. \n Continue your hard work, \n and you will become the best student at the School of Witchcraft and Wizardry!", 7

Object.create 1, 3, 0, 0 # gate
Object.create 6, 4, 2, 4 # magic stone
Object.create 19, 1, 6, 1 # tree base  
Object.create 21, 3, 7, 9 # tree
Object.create 22, 1, 4, 12 # yellow flower
Object.create 22, 2, 8, 6 # yellow flower
Object.create 25, 0, 13, 8#grass
Object.create 9, 0, 10, 0 #
Object.create 9, 2, 7, 0 #
Object.create 9, 3, 13, 0 #


Event.create "plarn05", 50, 2, 4


# answer: 
# Hero.movedown 3
#Hero.moveleft 5
#Hero.movedown 4
#Hero.moveleft 4
#hero.moveup 4
