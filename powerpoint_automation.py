# step : 1
# install python in your machine .


# step : 2
# pip install py32win


# step : 3
# create a powerpoint file make sure your keep powerpoint open


# step : 4
# create a python file

# step : 5
# import your library like this 
import win32com.client


# step : 6
PPTApp = win32com.client.GetActiveObject("PowerPoint.Application")
PPTPresentation = PPTApp.ActivePresentation


# step : 7
# how to duplicate a slide
PPTPresentation.Slides(1).Duplicate()


# step : 8
# how to delete a slide
PPTPresentation.Slides(2).Delete()


# step : 8
shapes  = PPTPresentation.Slides(1).Shapes


# step : 9
# loop thought all the shapes in the slide
for shape  in shapes:

	print(shape.Name)

	if shape.Name == "Title 1" :
		shape.TextFrame.TextRange.Text = "This is my header"


	if shape.Name == "Title 2" :
		shape.TextFrame.TextRange.Text = "This is my subtitle"





