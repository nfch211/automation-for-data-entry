There, is a workplace where i tried to automate repeatable tasks
for the company i worked for.

Most of the jobs was really simple data entry,
for example, entries of business registration numbers in excel
where the company names, business registration number, expiry dates
and company addresses will be filled in.

At the beginning, i asked ChatGPT for python script to 
read the pdf, and extract data to excel worksheet. 
It could be really done without understanding any fundamentals of 
programming language, in my case, i was using python.

Part A. Extracting information from image files

The algorithm of how it works is that I converted the pdf into docx,
named these files as same as the pdf. Then, i extracted all the information as mentioned from the docx,
saved these information to an excel file by a python script, 
with its' file names i could track back and link these information to the pdfs.
Finally, whenever i needed to input these information as requested,
all I need to do was simply to copy and paste.

Part B. Paste the information as according to my specifications to the excel file

I asked ChatGPT to make me another python script to automate the pasting process,
it was done by simply providing the column names and row numbers,
and ran the python script and then it was done.

Part C. Limitations

The quantity of pdf could range from 100 to 2000.
The processing time could range from 30 mins to 4 hours.
Issues were that, 

texts converted from pdfs were not 100% accurate.

texts in lines of the docxs were not structured or formatted like the image itself.

since it just extracted text, it failed to verify whether it is a valid business registration certificate,
for an example, a Hong Kong Business Registration Certificate.








