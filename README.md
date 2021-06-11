# Project_0001
My first project in VBnet

I am trying to resize a frame (via a class who inherit PictureBox) in a PictureBox containig an Image, but I have few issues with the method OnResize.

1- I can move this frame (red rectangle) via OnMove method (no problem)

2- I can resize this frame only with the corner bottom-right, which keep the ratio of the frame to 1.5 (landscape) or 0.66 (Portrait)

3- When I am resizing the frame, the resizing action should stop when it touch the right or bottom side, but it is working only partially.

4- To see the issue you have to move the red frame close to the right side or the bottom side and then try to increase the frame.

I tried various code options, but I can't find the solution.
Here you have a short version/application of what I am doing with the issue.

Any ideas to make progress are obviously welcome.

Thanks
