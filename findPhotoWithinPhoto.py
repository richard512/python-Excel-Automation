'''
Uses OpenCV to find what part of an image (large_image.png) contains another image (small_image.png)
Yellow Circle = Center of the location of small_image.png within large_image.png. The spot to click.
'''


import cv2
'''
if cv2.__version__[:1] == '2':
	import cv2 as cv
	method = cv2.TM_SQDIFF_NORMED
elif cv2.__version__[:1] == '3':
'''

#import cv2 as cv
method = cv2.TM_SQDIFF_NORMED

# Read the images from the file
small_image = cv2.imread('small_image.png')
large_image = cv2.imread('large_image.png')

result = cv2.matchTemplate(small_image, large_image, method)

# We want the minimum squared difference
mn,_,mnLoc,_ = cv2.minMaxLoc(result)

# Draw the rectangle:
# Extract the coordinates of our best match
MPx,MPy = mnLoc

# Step 2: Get the size of the template. This is the same size as the match.
trows,tcols = small_image.shape[:2]

# TOP LEFT circle
center = (MPx,MPy)
radius = 5
color = (0,255,0)
cv2.circle(large_image, center, radius, color, thickness=1, lineType=8, shift=0)

# BOTTOM RIGHT circle
center = (MPx+tcols,MPy+trows)
radius = 5
color = (0,255,0)
cv2.circle(large_image, center, radius, color, thickness=1, lineType=8, shift=0)

# CENTER circle
center = (MPx+tcols/2,MPy+trows/2)
radius = 5
color = (0,255,255)
cv2.circle(large_image, center, radius, color, thickness=10, lineType=8, shift=0)

# Step 3: Draw the rectangle on large_image
cv2.rectangle(large_image, (MPx,MPy),(MPx+tcols,MPy+trows),(0,0,255),2)

# Display the original image with the rectangle around the match.
cv2.imshow('output',large_image)

# The image is only displayed if we call this
cv2.waitKey(0)
