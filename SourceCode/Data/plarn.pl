Game.autorun

# clear everything
Object.clear
Pathway.clear
Message.clear
Block.clear
Event.clear
Pathway.block 0

a = 3 # declare a variable
Tile.style a # types of floor tiles  
Pathway.clear # clear up previous variables  
Pathway.style a # grassplot setting


Object.create 8, 2, 10, 6
Event.create "plarn01", 50, 10, 6


Message.clear
Message.show "Magic World! ", 2, 8, 0, 9
Message.show "Level 1 \n Hi there! Welcome to the Magic World. \n Python has special power here \n and can be used as spells to control this world! \n Mission: \n Control your hero to move around \n and get the hammer! \n hint: hero.moveright ", 6
Hero.say "I need to get the hammer", 1

Hero.create "Harry Potter", 6, 6  
Hero.steps 2, 2, 2, 2  # auxiliary instruction tool For demonstrating the wizard ' s walking route  


