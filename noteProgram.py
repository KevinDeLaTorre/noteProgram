# Author : Kevin De La Torre
# Purpose: Take a ".notes" file and convert it to a powerpoint file using "python-pptx"

from sys import argv
import os
from pptx import Presentation

PRES_FOLDER_NAME = "presentations"

def main():
    if ( len( argv ) > 1 ):
        folderPath = os.getcwd() + '/' + PRES_FOLDER_NAME
        if not os.path.exists( folderPath ):
                os.makedirs( folderPath )

        for i in range( 1, len( argv ) ): # Go through all files
            file = open( argv[ i ], 'r' ) # Open file in read-only mode
            prs = Presentation()
            for line in file:
                line = line.rstrip()            # Chop trailing newline
                print( line )

            # Gets rid of folder name at left then replace file extension and places in folder
            prs.save( folderPath + '/' + argv[ i ].split( '/', 1 )[ 1 ].rsplit( '.', 1 )[ 0 ] + ".pptx" ) 
            file.close()
    else:
        print( "Usage: {0} <file1> <file2> ... <fileN>".format( argv[ 0 ] ) )

if __name__ == "__main__":
    main()
