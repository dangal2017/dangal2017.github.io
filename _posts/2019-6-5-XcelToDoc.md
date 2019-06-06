---
layout: post
title: Age old Excel to Doc converter but OFFLINE!
---



![_config.yml]({{ site.baseurl }}/images/meraCanadianKumar.gif)


Offline? Seriously?  

If someone asks you that they want to convert an excel document to a doc version, they will tell you about the infinite converters available online.  

But, what if you don't want to share your excel data to an online site?  

Then XcelToDoc does all the same conversion but offline. Sounds all rosy, but it doesn't actually does it without any manual effort like those sophisticated online ones. :(   

It requires manual effort in adding tags in your input excel telling where a para or a heading or a table starts.  

So, before start of every paragraph or a heading, you need to give a tag "P" or "H" respectively in the cell before the paragraph or heading, telling our script that there is a paragraph or heading in the next cell. It will format them accordingly.  

Currently, the formatting for a paragraph and a heading is fixed at:  

Font - Arial  
Font size - 10 for a paragraph and 14 for a heading  

You can change them in the script very easily.  

A table however isn't this simple. You need to indicate 3 tags:  

"TS" - indicating start of a table, in the cell before the start of first row  
"TRE" - indicating end of the rows of a table, in the cell before the start of last row  
"TS" - indicating end of the columns of a table, in the cell after the end of first column  

Find a dummy excel file xcelToDoc.xlsx in the github repository, illustrating how to add these tags.  

Also, just like the the formatting for paras and headings, table formatting is currently fixed at:  

Font - Arial   
Font size - 10  
First row of table - Bold  

Again, you can change them in the script very easily!    

Auto width for a table cell is set to True.  

Even though autofit fits to page width for a few tables but for most tables it doesn't. Need to fix this. It will be fixed in a later version.  

Also, don't forget to remove any formulas, currently it prints what the formula is, though it should fill the value in the cell after the formula is applied. It will be fixed in next version.    
  
Thus, the steps to using xcelToDoc are as follows:  

Step : 1 Make the input excel in a format that our script understands, as illustrated above  

Step : 2 Put all XcelToDoc files and input excel file in a single folder.  

Step : 3 Run the script XcelToDoc from the above folder created in Step 2  

If you don't have python installed on your system, run exe file by double clicking it.  

If you have python installed, go to terminal and navigate to the folder, and run the script by typing the command -
python \<script name\>  

If you want to run the script through jupyter notebook, open the file in jupyter notebook and press "Run all" button at the top.  

Step : 4 A UI opens after you run the script in step 2. Input your master excel name in first box and then input the name of the updated doc that you want.  

Step : 5 Click on "Get updated doc" button to get updated doc. The updated doc gets stored in the same folder as created in step 2.  

Pay heed to the exception messages printed on the UI. They are important to catch any errors that occurred during the conversion.  

As with every software or script, this script might have bugs which are not yet reported or found. Please Quality check your work twice before sending out to any one else. This script is protected under the MIT license.    

MIT/X11 is basically a simple contract that says:  

Person or company X created Y.  
Y belongs to X, but X is granting you the right to use it and do whatever you want with it.  
X cannot be held accountable for anything that goes downhill with what you do with Y.  

Report any bugs you face at the github page. Feel free to use the code at:  
https://github.com/dangal2017/XcelToDoc
