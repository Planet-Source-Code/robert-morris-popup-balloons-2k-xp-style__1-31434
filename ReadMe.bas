Attribute VB_Name = "ReadMe"
  '***********************************************************************
  '## THIS MODULE *IS NOT* NEEDED FOR DOING ANYTHING IN THE PROJECT, SO ##
  '## IF YOU MAKE A PROJECT OF YOUR OWN, YOU *DO NOT* NEED TO INCLUDE   ##
  '## THIS MODULE, README.MOD.                                          ##
  '***********************************************************************

'VB Popup Balloons (Release 1.0, Feb. 2, 2002)
'By, Robert Morris
'http://rmsoft.itgo.com
'
'THIS IS THE SAME DOCUMENT AS README.TXT -- I JUST PUT IT AS AN ENTIRELY
'COMMENTED MODULE IN THE PROJECT.
'
'This sample VB project shows you how to use Windows 2000/XP-style popup
'balloons in your programs. Microsoft introduced balloons in Windows 2000,
'and they're still there and used in Windows XP--especially from the
'system tray, which is where they're mostly seen from. However, they are
'also used in various programs--such as the PowerToy Calculator if you enter
'something wrong.
'
'Well ... Microsoft hasn't released how to show these cool balloons in your
'programs yet. (They have for tray icons, but that's not what we're doing
'here. This is for forms, or whever your creativity takes you!) So, instead
'of waiting or trying to dig through API myself, I just decided to create
'a way of showing balloons myself. Besides, the "real" ones will only work
'on Windows 2000 or later. This will work on anything since Windows 95.
'
'Take a look at the sample project's examples, and you'll quickly see what
'uses they have. They can be used as a sort of non-interrupting message
'box--that is, for something you need to tell the user about but don't
'really need to interrupt him/her with a message box about, as is done in
'the sample division program included here upon a division by zero or other
'mathematical error.
'
'These balloons support many features. Like a messagebox, you can have
'them show an "i", "x", or "!" icon (in the upper left). No question mark,
'though, since I figured you can't ask a question using one of these!
'You can also define the text (of course!) to be shown, the bold title above
'the text, the font, and the position of the balloon. Additionally, you
'can have the balloon automatically close after a specified amount of time,
'and you can control whether or not an "X" close button is displayed on the
'balloon. Clicking anywhere in the balloon will make it disappear. There's
'more! And since the full source is included, you can literally customize
'it any way you want.
'
'Comments throughout the code tell you what it's doing and how you can
'use it, but here's an overview.
'
'The form frmBalloon is a "template" form. We won't use it directly; every
'time we want a balloon, we'll create a new instance of it, as in:
'   Dim frmMyBubble As New frmBalloon
'
'Then, we need to set the properties for the balloon--text, icons, etc. To
'do this, we call the form's (frmMyBubble, in this example) SetBalloon sub,
'as in:
'   frmMyBubble.SetBalloon "Title Here", "Text Here", 140, 230 ...
'that 's just an example. See the code's comments (frmBalloon, go to (General),
'and go to SetBalloon). The dots at the end just indicate that you should
'go on; don't use them in the code. Actually, above are all the REQUIRED
'arguments to pass to the sub, but you will probably want to pass more.
'
'Notice the 140 and 230 in the example above. Those are the screen
'coordinates you want the bubble to be shown at. Since you most likely
'do not know the exact location in pixels and since you probably want
'it shown by another control on your form, you'll need to use some API
'to determine just where that control is on the SCREEN (it's Top and Left
'properties are relative to the form, so they won't do) and then pass that
'to SetBalloon.
'
'See the example for how it's done--getting the coordinates. It's on
'frmSample, in cmdPopIt_Click(). It's commented quite a bit.
'
'After you set the balloon's properties, you can then show it, as in:
'    frmMyBubble.Show , Me
'The " , Me" will set whatever form you're calling this from as the owner
'form of this balloon. You'll want to do this. And, we also need to add
'this right after we show the balloon:
'    Me.SetFocus
'which will give the form you showed the balloon from the focus again.
'Showing the balloon will take focus away from the form, which we don't
'want to happen. This works around that until I can find the API for showing
'a window (the balloon, in this case) without "stealing" focus.
'
'What you don't get from reading this you should find in the code -- enjoy!
'
'Also, I am aware that this isn't exactly perfect yet, especially when you
'adjust the size of the balloon. When you call SetBalloon, you'll see the
'default properties for lHeight and lWidth. They're a pretty normal size,
'and if you want to change them it would be best not to go too far away
'from that size, otherwise the balloon's border (a Shape control) will not
'fit Right.I 'm trying to work on that; the code I'm using to round the
'corners of the balloon doesn't fit exactly with the shape of the shape
'control.
'
'Feel free to modify this code and/or use it in your own projects, but
'please don 't post the pure source at a source code site--that is, don't
'take credit for my work. However, if you have modified and improved the
'source in some way, then feel free to post it somewhere--just drop me
'an e-mail first. Thanks!
'
'---------------------------------------------------------------------
'
'Robert Morris
'http://rmsoft.itgo.com
'robertmorris@ softhome.net
'
'-THIS SOURCE CODE WAS POSTED AT PLANET-SOURCE-CODE.COM AND IN THE AOL
'VISUAL BASIC LIBRARIRES. PLEASE DO NOT POST IT ANYWHERE ELSE EXCEPT AS
'DEFINED ABOVE. THANKS-
