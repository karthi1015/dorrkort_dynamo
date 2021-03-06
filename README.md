# Using an excel sheet as dictionary in Dynamo

## Where to start
**Some of this text might not be interesting to everyone. If you're just looking for a solution to import Excel sheets as dictionaries into Dynamo you can jump ahead or download the dyn-file right away. But for Swedish professionals with the same problem or someone who loves to read, the rest gives you some context.**

## Dependencies
Testet with Dynamo 2.x and Revit 2020. All nodes are OOTB!

## The "Dörrkort"-problem
The office I'm working in is switching from a long legacy in in AutoCAD Architecture over to Revit. There are some test projects and I'm part in one of them. The plan is to create a smooth workflow that enables us to take advantage of the built in solutions but also exended functionality of Revit and BIM. 
A very tidious job when building here in Sweden is the so called "Dörrkort". That's basically one door per drawing and you have all door information sorted in a standardized manor. The only good example I could find was in [this master thesis](http://www.diva-portal.se/smash/get/diva2:1221728/FULLTEXT01.pdf) on page 77. Our current workflow in AutoCAD demands a lot of manual checking and filling out text in the respective row. That's just how it was handled so far, there might be a better solution for AutoCAD as well, but my focus is now on Revit.
The biggest part is definitely to define door specific locks and acessories like automatic openers and more. It feels like it can't be so difficult, but someone actually needs to define them. In my current project I have ~180 doors with 27 different lock/acessory-combinations. There are obviously projects with even more doors, that's why I want to automate as much as possible.

## Theory
Lets imagine you have a Excel list from a consultant with different features that apply to more than just one instance in your Revit project. I will stick with my example, 10 doors with 5 different lock types, but each door has a lock meaning that some locks share the same lock type. That's a perfect scenario for a dictionary.
This could be the the excel sheet, in reality there are more columns and more rows. DK is the key for the dictionary.
|DK|Material|Lock|Handle|
|-|-|-|-|
|DK1|Steel|A|round|
|DK2|Steel|B|straight|
|DK3|Steel|C|straight|
|DK4|Aluminum|D|straight|
|DK5|Steel|E|straight|

There are 10 doors in the project:
|ID|DK|Material|Lock|Handle|
|-|-|-|-|-|
|00|DK5|-|-|-|
|01|DK2|-|-|-|
|02|DK3|-|-|-|
|03|DK1|-|-|-|
|04|DK5|-|-|-|
|05|DK4|-|-|-|
|06|DK2|-|-|-|
|07|DK1|-|-|-|
|08|DK1|-|-|-|
|09|DK4|-|-|-|

In the end the list should look like this. Again, imagine that for 180 doors, more parameters and surely more complex names. You should also be aware that the doors might change over time, it's a nightmare to do that by hand:
|ID|DK|Material|Lock|Handle|
|-|-|-|-|-|
|00|DK5|Steel|E|straight|
|01|DK2|Steel|B|straight|
|02|DK3|Steel|C|straight|
|03|DK1|Steel|A|round|
|04|DK5|Steel|E|straight|
|05|DK4|Aluminum|D|straight|
|06|DK2|Steel|B|straight|
|07|DK1|Steel|A|round|
|08|DK1|Steel|A|round|
|09|DK4|Aluminum|D|straight|

## What Dynamo does
Dynamo can't apply a dictionary from an excel sheet and just match the key value of the excel sheet with the key value of Revit objects out of the box (afaik). You basically have to tell Dynamo to get all the doors, and then you match each parameter in the door to the parameter in the excel sheet/dictionary. That way you get a long list. I would say it's a workaround, there should be a node that does that out of the box.

Here is what I did:
![Dynamo Script](screenshots/complete.png)

## Preparations
To have the script working you need to create parameters in your project. I used shared parameters that I loaded as project parameters and in this case most of them are instance-based. But how you handle your parameters is your thing, the important part is that this script doesn't create the parameters, you have to have them beforehand! If you google a bit you'll find tips how to batch create parameters.

## Walk through the script
### Step 1 - loading Revit elements
![Step 1](screenshots/step_1.png)

That's quite basic, but worth mentioning when using Dynamo for the first time. You can use the script for any other category.

### Step 2 - loading and processing the Excel sheet
![Step 2](screenshots/step_2.png)

Excel import has it's own node, that's quite easy. Then I use two Python scripts, it might be possible without python, but it's faster for me to write a python script. Let's follow the upper line first. It takes the headers of the Excel sheet, Those are the parameter names that get filled out later - it's important that there are no typos! Then I drop the key row as I only want to have the parameters that I want to add to the doors, the key stays unchanged.

In the very end you can see I change the names from ALL CAPS to title, so that the first letter of each word is in capitals. If you take better care of your parameters that's not necessary, but be aware that Dynamo is case sensitive!
```python
a = IN[0]
b = []
# Place your code below this line
b = [i.title() for i in a]	

# Assign your output to the OUT variable.
OUT = b
```
On the lower line I start with a python script that basically creates the dictionary:
```python
data = IN[0]
dk = [i[0] for i in data]
params = [i[1:] for i in data]

di = dict(zip(dk,params))

OUT = di
```
And then I use the OOTB-node Dictionary.ValueAtKey to read all values for all the door-parameters I want to change (the ones in the string above). The key for the values is the instance parameter "BeslagNummer" (DK) in the project.

### Step 3 and 4 - crucial!
What now happens is the crucial part (which Dynamo should have as a node because it doesn't seem like a lot of nodes...).

![Step 3](screenshots/step_3.png)

What Dynamo now does is basically calling each element as often as there are keys in the dictionary. On the upper line you see that it reads the lenght of the keys list and then it repeats all the elements from the door cateogory by the key amount. Lets say you have 30 keys and 200 elements that means that you created a list of 6000 elements that you want to write in the next step. Horrible to do that by hand!
The list looks something like this if we use the example from above with 10 doors and 3 parameters to write:
`[01,01,01,02,02,02,03,03,03,04,04,04,05,05,05...]`

In the lower line it counts the amount the elements and repeats the list of dictionary keys as often as there are elements, with the example from above:
`[Material,Lock,Handle,Material,Lock,Handle,Material,Lock,Handle,Material,Lock,Handle,...]`

Now everything comes together and suddenly makes sense. Dynamo runs through the script and loops through each key/parameter in the excel sheet on each door in the revit project. In my case it was around 5800 writings.

### Thanks
This was my first real life Dynamo script and I have to thank [the guys at the dynamo forum](https://forum.dynamobim.com/t/using-dictionaries-with-excel-and-parameters/57182) who helped me with it! I hope someone can make use of it and feel free to propose improvements by forking or creating a pull request.
