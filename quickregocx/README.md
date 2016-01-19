# VB6-Applications
VB 6 applications I've made about 10-15 years ago

#Quick Reg OCX control

 Have you ever wanted to have your project set up to have people have to register with you to unlock certain parts of the project. Well, i have made a simple to use control just for this purpose. Just slap this control onto a form, set a few values, and you will be able to generate custom keys depending on the name of the users computer, or a serial number. You can also use a combonation of these to make a harder to crack key. What the code does is, if the setting is set to ComputerName, it finds out the name of the computer, goes through my alogorithim, and makes a key for that name. When the use puts that key into a text box or something, you can call my control to see if that key is correct. The control will respond with a boolean value, if the key is correct, it will be true, but if the control determines the key to not be right, it will reject the key with a false value. This code is heavily commented (prolly more comments then code), but the code is effectivly explained (well in my eyes it is), and will show you how to set this up. Also included in this visual basic group, is a sampel program giving the basics on how to use the control. 

 Please do not use this as it is not secure. ;)