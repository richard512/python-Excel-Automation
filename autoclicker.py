'''
Automatic Mouse Click based on screen contents
Waits until the pixel at coordinates (364, 278) have the RGB color (75, 219, 106), then clicks
'''

import gtk.gdk
import sys
from pymouse import PyMouse
from time import sleep

m = PyMouse()

def PixelAt(x, y):
    w = gtk.gdk.get_default_root_window()
    sz = w.get_size()
    pb = gtk.gdk.Pixbuf(gtk.gdk.COLORSPACE_RGB,False,8,sz[0],sz[1])
    pb = pb.get_from_drawable(w,w.get_colormap(),0,0,0,0,sz[0],sz[1])
    pixel_array = pb.get_pixels_array()
    return pixel_array[y][x]

def main():
	while True:
		x = m.position()[0]
		y = m.position()[1]
		r, g, b = PixelAt(x, y)
		print m.position(), " == ", r, g, b

		r, g, b = PixelAt(364, 278)
		if (r, g, b) == (75, 219, 106):
			print "green = click!"
			m.click(400, 223)

main()
