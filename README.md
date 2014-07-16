ComAddinEvent
=============

Proof of concept for communication between COM addins within VSTO context

To debug:

1. Open and run HasEvent. Word should start and you should see the add-in load.
2. Close Word. The add-in HasEvent is still installed, and will load next time you run Word.
3. Open ConsumeEvent
4. Reference HasEvent.dll in ConsumeEvent. (References → Add reference → lookup HasEvent.dll, usually in HasEvent/bin/Debug)
5. Run ConsumeEvent. Word should start and you should se HasEvent and ConsumeEvent load.
6. Open or New a document, then attempt to save it. You should get a messagebox.
