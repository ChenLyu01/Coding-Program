Game.autorun

Tile.style 2
Pathway.style 2

#======= clear everything=======
Object.clear
Pathway.clear
Message.clear
Block.clear
Event.clear
Pathway.block 1
# =========================

# ========Create Player=========
Hero.create "Harry Potter", 2, 8  
Event.create 7, 2, 8
Message.show "You are wonderful! I admire your hard work. You will be the best student in the school of magic and wizardry. ", 7
# =========================

Hero.steps  2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3,3,3,3
Pathway.draw 0, 0, 2, 14  
Pathway.draw 3, 0, 8, 4
Pathway.draw 2, 9, 1, 1
Pathway.draw 3, 4, 3, 4
Pathway.draw 3, 9, 3, 5
Pathway.draw 7, 5, 4, 3
Pathway.draw 7, 9, 9, 5
Pathway.draw 12, 7, 4, 2
Pathway.draw 12, 0, 3, 6
Pathway.draw 15, 0, 1, 4
# Hero.say "type: \n hero.moveright", 1
Hero.say "I am a hero!", 1

Message.show "Task: \n Please type: hero.moveright \n hero.moveup", 6
# Message.show "Task: \n Please type: hero.moveright \n hero.moveup n", 1


Object.create 1, 3, 0, 0 # gate
Object.create 1, 3, 0, 0 # gate
Object.create 19, 1, 6, 1 # tree base  
Object.create 21, 3, 7, 9 # tree
Object.create 22, 1, 4, 12 # yellow flower
Object.create 22, 2, 8, 6 # yellow flower
Object.create 25, 0, 13, 8 # grass
Object.create 9, 0, 10, 0 #
Object.create 9, 2, 7, 0 #
Object.create 9, 3, 13, 0 #
# Object.create 7, 0, 2, 5 #


Object.create 6, 3, 11, 4 # map
Event.create "plarn03", 50, 11, 4




