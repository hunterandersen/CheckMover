# CheckMover
A simple script I wrote at my last job. I made it specifically to expedite one task I had of renaming and moving pdfs from one folder to another folder.

## Description
At my previous job, I would receive scanned PDFs of various checks from our company's vendors. Because the checks were scanned, none of the file names contained the one piece of information we (humans) cared about - the check number. This means that for every check, I had to 
1. Open the PDF
2. Read the check number
3. Close the PDF (because you can't rename a file that is open)
4. Rename the file to the check number
5. Move the file to the corresponding vendor folder on our company's file system.

As you may imagine, depending on the number of files to move, this task could vary from simply tedious, to rather cumbersome. Once I wrote this script it drastically reduced the amount of time required to perform this rather simple task.

## Made using [AutoIt](https://www.autoitscript.com/site/autoit)

I found AutoIt after some google searching. I wanted something simple to pick up that would let me automate a very mundane task in Windows. This fit the bill perfectly, and so I gave it a try. I'd never used it before or even heard of it before, but I'm not easily intimidated by new programming languages.
It served my purposes well, though I do wonder if Python may have been a better choice long-term. Oh well, I enjoy this little script I made.
