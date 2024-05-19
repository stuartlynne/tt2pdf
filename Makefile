# vim: shiftwidth=8 tabstop=8 noexpandtab textwidth=0

#

tt2pdf-onefile:
	time -p python -m nuitka \
		--onefile \
	 	--windows-disable-console \
		--disable-plugin=pyside6 \
		--disable-plugin=tk-inter \
		--noinclude-unittest-mode=nofollow \
    		--windows-icon-from-ico=./images/rfid_5518063.png \
		tt2pdf.py
	#belcarra-signtool tt2pdf.exe

tt2pdf-console:
	time -p python -m nuitka \
		--onefile \
		--disable-plugin=pyside6 \
		--disable-plugin=tk-inter \
		--noinclude-unittest-mode=nofollow \
    		--windows-icon-from-ico=./images/rfid_5518063.png \
		tt2pdf.py
	#belcarra-signtool tt2pdf.exe
