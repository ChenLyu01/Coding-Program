game.autorun

a = 1
tile.style a
pathway.style 1
pathway.clear
pathway.draw 8, 8, 3, 3

for a = 1 to 3
    pathway.draw 1, 1, a, 7
next 
message.show "Hello world! ", 0 , 0 ,0 ,9
message.clear
message.show "Hello, Welcome to the magic world! ", 0 
hero.create "Harry Potter", 1,1 
hero.moveleft (3) 
hero.moveright 5
hero.moveup 4
hero.movedown (4)
block.show
block.hide
event.show
event.create 61, 4, 4
event.destroy 4, 4

object.clear
object.show
object.create 1, 2, 4, 4
object.destroy 1, 2

block.create 4, 4
object.destroy 4, 4

effect.create 1, 1, 4, 4 
effect.enable 1, 1
effect.disable 1, 1

event.create "do_loop_until", 50, 4, 4

'file.open "do_loop_until"