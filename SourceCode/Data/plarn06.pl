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

Event.create "plarn07", 50, 10, 4

Hero.say  "I need to collect the Knotgrass.", 1
Message.show "Level 7 \n Mission: Open the door and collect Knotgrass \n Hint: \n hero.faceup \n hero.open \n hero.moveleft \n hero.movedown \n hero.moveup", 6
Message.show "You can see the Knotgrass there. To open the door, you need to face the door and then open it. ", 7
Hero.create "Mike", 10, 5
Event.create 7, 10, 5

Hero.steps 3, 3, 3, 3 


Object.create 25, 0, 1, 0
Object.create 25, 1, 0, 1
Object.create 19, 0, 6, 9
Object.create 24, 0, 4, 1
Object.create 25, 2, 11, 6

Object.create 1, 0, 0, 3
Object.create 9, 0, 9, 4
Object.create 25, 0, 7, 1
Object.create 25, 1, 9, 2
Object.create 25, 2, 11, 1 # door
Object.create 25, 3, 12, 1
Object.create 29, 0, 13, 11 # three color flower
Object.create 18, 4, 13, 10  
Object.create 18, 5, 12, 10  
Object.create 28, 0, 2, 1
Object.create 24, 1, 1, 1


Pathway.draw 0,0,16,14


a = 0
b = 5
c = 0

For a = 8 To 13
        Pathway.destroy 1, a
Next  

For a = 7 To 13
        Pathway.destroy 5, a
Next  

For a = 5 To 7
        Pathway.destroy 10, a
Next

For a = 9 To 10
        Pathway.destroy 10, a
Next  
For a = 7 To 9
        Pathway.destroy 14, a
Next  

For a = 1 To 10
        Pathway.destroy a, 11
Next  

For a = 1 To 10
        Pathway.destroy a, 11
Next  

For a = 5 To 10
        Pathway.destroy a, 7
Next  

For a = 10 To 14
        Pathway.destroy a, 9
Next

For a = 14 To 15
        Pathway.destroy a, 7
Next  

For a = 6 To 14
    For b = 0 To 4
        Pathway.destroy a, b
    Next  
Next
