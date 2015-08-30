#!/usr/bin/pyhton

#utility to move and copy files. files are hidden during actual copy
#april 24 - 2014 - †
#armen karamian

import os, shutil as sh, stat

def fileMover(source, destination, mode):	#file mover #get incoming file #set destination path set mode
	if not (os.path.exists(source)):
		raise Exception(source+" does not exist")
	if not (os.path.exists(destination)):
		raise Exception(destination+" does not exist")
	filename = os.path.split(source)[1]
	hiddename = '.'+ filename															#file will be hidden during copy/move
	hiddenfinaldest = os.path.join(destination, hiddename)

	if (mode == 'copy'):				#copy/move
		sh.copy2(source, hiddenfinaldest)
		finaldest = os.path.join(destination,filename)
		os.rename(hiddenfinaldest, finaldest)
		return finaldest
	if (mode == 'move'):
		sh.move(source, hiddenfinaldest)
		finaldest = os.path.join(destination,filename)
		os.rename(hiddenfinaldest, finaldest)
		return finaldest
	else:
		raise Exception("Mode not specified (copy/move)")
	
	
	
	
if __name__ == "__main__":
	src = '/Volumes/fs3//encoding/AssetManagement/TestingNBC/Suts.mov'
	dest = '/Volumes/fs3//encoding/AssetManagement/TestingNBC/testdir'
	print fileMover(src,dest, 'move')
		