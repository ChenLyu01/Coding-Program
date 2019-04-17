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

Hero.create "Mike", 0, 9
Event.create 7, 0, 9

Hero.steps 3, 3, 3, 3, 2, 2, 2, 0, 0, 0, 2, 2, 2, 3, 3, 3, 2, 2, 2, 0, 0, 0, 2, 2, 2, 3, 3, 3, 2, 2, 2, 0, 0
hero.say "I am the best!", 1
Message.show "Task: \n Please type: \n Hero.movexy X, Y or type: \n  a = 0 \n  for a = 0 to 2  \n       hero.moveup 3 \n       hero.moveright 3 \n       hero.movedown 3  \n       hero.moveright 3 \n next ", 6
Message.show "Task: \n Please type: \n Hero.movexy X, Y or type: \n  a = 0 \n  for a = 0 to 2  \n       hero.moveup 3 \n       hero.moveright 3 \n       hero.movedown 3  \n       hero.moveright 3 \n next ", 7

event.create "plarn01", 50, 15, 9
Object.create 6, 0, 15, 9

Pathway.draw 0, 0, 16, 14
Object.create 21, 0, 10, 1.
Object.create 24, 5, 14, 4
Object.create 28, 0, 8, 4
Object.create 24, 0, 9, 4
Object.create 24, 1, 5, 10
Object.create 24, 2, 6, 12
Object.create 22, 3, 7, 11
Object.create 24, 4, 8, 12
Object.create 22, 5, 3, 11
Object.create 24, 6, 0, 11
Object.create 28, 0, 1, 2
Object.create 19, 2, 4, 2
Object.create 28, 3, 5, 4
Object.create 28, 4, 6, 1
a = 0
b = 0
c = 0
d = 0

For b = 0 To 5
    c = b * 3
For a = 6 To 9
        Pathway.destroy c, a
Next  
Next  

a = 6
For d = 0 To 3
        Pathway.destroy d, a
Next  

a = 6
For d = 6 To 8
        Pathway.destroy d, a
Next  

a = 6
For d = 12 To 14
        Pathway.destroy d, a
Next  


a = 9
For d = 4 To 6
        Pathway.destroy d, a
Next  

a = 9
For d = 10 To 12
        Pathway.destroy d, a
Next  

a = 0
For d = 0 To 3
        Pathway.destroy d, a
Next


