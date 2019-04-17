game.autorun

Hero.create "Harry Potter", 2, 3  

Tile.style 2
Pathway.style 2

# Clear everything
Pathway.clear
Pathway.block 1
Message.clear
block.clear

Hero.steps 0, 0, 0, 0, 0
Pathway.draw 0, 0, 2, 14  
Pathway.draw 3, 0, 8, 4

Pathway.draw 3, 4, 3, 4
Pathway.draw 3, 9, 3, 5
Pathway.draw 7, 5, 4, 3
Pathway.draw 7, 9, 9, 5
Pathway.draw 12, 7, 4, 2
Pathway.draw 12, 0, 3, 6
Pathway.draw 15, 0, 1, 4
Hero.say  "I need to collect magic stones.", 1

message.show "I am a special magic book. I can help you become a magician. ", 1
message.show "Congratulations! you mastered the most basic spell in this Magic World. You have become the assistant of Professor Vold who can make very valuable potions. He just asked you to collect some plant ingredients to make a magic potion! ", 7
message.show "Level 2 \n Mission: \n Get the magic stone to get the map of the Dark Forest full of plant ingredients! \n Hint: hero.movedown", 6

Object.clear
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
Object.create 7, 0, 2, 5 #


Object.create 6, 3, 2, 8 # map
Event.create "plarn02", 50, 2, 8
Event.create 1, 2, 5
Event.create 7, 2, 3


