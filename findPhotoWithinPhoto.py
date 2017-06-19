'''
Uses OpenCV to find what part of an image (large_image.png) contains another image (small_image.png)
Yellow Circle = Center of the location of small_image.png within large_image.png. The spot to click.
'''

import cv2
method = cv2.TM_SQDIFF_NORMED

# get the images
small_image = cv2.imread('small_image.png')
large_image = cv2.imread('large_image.png')

# look for small_image inside large_image
result = cv2.matchTemplate(small_image, large_image, method)

# minimum squared difference
mn,_,mnLoc,_ = cv2.minMaxLoc(result)

# Coordinates of our best match
MPx,MPy = mnLoc

# Get the size of the template. Same size as the match.
trows,tcols = small_image.shape[:2]

# draw TOP LEFT circle
center = (MPx,MPy)
radius = 5
color = (0,255,0)
cv2.circle(large_image, center, radius, color, thickness=1, lineType=8, shift=0)

# draw BOTTOM RIGHT circle
center = (MPx+tcols,MPy+trows)
radius = 5
color = (0,255,0)
cv2.circle(large_image, center, radius, color, thickness=1, lineType=8, shift=0)

# draw CENTER circle
center = (MPx+tcols//2,MPy+trows//2)
radius = 5
color = (0,255,255)
cv2.circle(large_image, center, radius, color, thickness=10, lineType=8, shift=0)

# draw red rectangle
cv2.rectangle(large_image, (MPx,MPy),(MPx+tcols,MPy+trows),(0,0,255),2)

# display the results
cv2.imshow('output',large_image)

# wait for key press with 0 delay
cv2.waitKey(0)
