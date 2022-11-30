# officeXXE

XXE payload generator for office files.

Author: GigaByteRex

## Intro
officeXXE creates office files (currently docx and xlsx) with embedded XXE payloads to test for XXE in XML parsers used to read these documents.
The idea is to insert your own callback monitoring address, for example a Burp Collaborator instance, in a XML entity SYSTEM call, and check whether the XML parser handling your document makes a outbound query towards said address.

This tool is hyper-early alpha version levels of work-in-progress, so expect bugs and issues!

## Features

Currently the only working attack vector is docx payload, which has been tested with some level of success on sample document parsers.
The xlsx payload attack does not currently create documents that successfully trigger XXE during testing, but has some placeholder functionality to build on for further development.
Clusterbomb attack for both docx and xlsx is currently not implemented, but is a planned feature to replace ALL links in the document with your callback address, which will result in a broken document normally, but might still trigger callbacks for some XML parsers.

## Usage

officeXXE can be run as a interactive wizard by giving no arguments.
Alternatively you can provide officeXXE with the desired filetype, attack type, callback domain and output file name to generate the file directly.

Install the dependencies:

    pip install -r requirements.txt

Running officeXXE as a interactive wizard:

	python officeXXE.py

Running officeXXE directly:

	python officeXXE.py [options]

Options:

		-h, --help				show this help message and exit
		-f FILETYPE, --filetype FILETYPE
							Output file type, 1 for docx or 2 for xlsx.
		-a ATTACK, --attack ATTACK
							Attack type, 1 for payload or 2 for clusterbomb.
		-d DOMAIN, --domain DOMAIN
							Your callback monitoring domain, 
							e.g. a Burp Collaborator instance.
		-o OUTPUT, --output OUTPUT
							Output file name, with file extension.

## Contribution

There is tons of room for improvement and further development as this is primarily a hacked together POC relying on duct tape to stay alive. :face_with_spiral_eyes:
Thus any pull requests or improvement ideas are highly appreciated!
The most urgent feature I see is changing the behaviour of payload generation. Currently it, for docx, keeps a boilerplate word/document.xml file with a custom XML entity tag hardcoded into it. This is sub-optimal, and it should instead read the document.xml file after creation and insert the XXE payload where appropriate, instead of overwriting the entire thing with boilerplate.
The procedure is the same for xlsx, by doing the same with a boilerplate xl/sharedStrings.xml file.
By making sure the XXE payload is inserted dynamically instead of a pure overwrite the tool could avoid a bunch of possible formatting errors, and gets rid of the hassle of maintaining the boilerplate text blobs.
