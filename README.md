# ViOrb 

Designed to be the last start button replacement tool for all versions of Windows

## Background

I wrote a lot of the original code a long time ago and it was for a completely seperate project, ViStart. It was messy and sloppy but as long as it worked that's all I initially cared about. I've sinced tried to tidy up the code to make it more presentable. In the process of doing that I also ended up recoding a lot of it and rediscovering how useful comments would have been. 

## Libraries

- [Windows Unicode API TypeLib](https://github.com/badcodes/vb6/blob/master/%5BInclude%5D/TypeLib/winu.tlb) - Windows API, stores all the API declerations
- [Karl E. Peterson's - HookMe](http://vb.mvps.org/samples/HookMe/) - A clean and elegant means of sublcassing 
- [vbAccelerator - GDIPlusWrapper](https://github.com/tannerhelland/vbAccelerator-Archive/tree/master/VB/Code/vbMedia/Using_GDI_Plus/GDIPlus_Helper) - vbAccelerator's GDIPlusWrapper used for OOP GDIPlus

## Getting Started 

Grab the WinU TLB - extract the TLB and add as a reference to the project
Grab the HookMe zip - extract the files (IHookSink.cls, MHookMe.bas) over the place holder files (IHookSink.cls, MHookMe.bas) and disregard any other files
Grab the GDIPlusWrapper zip - extract contents to "GDIPlusWrapper" 

Add the "Clear" function from the Prototype folder to the GDIPlusGraphics class then compile it and add the resulting binary to ViOrb
Release/GDIPlusWrapper.dll

## Comments

I rarely get time to update these projects so I am releasing them here in hopes they can be useful to someone :) I don't expect anyone would care to continue development. I am not against them doing so. If I ever got time I would have liked to have recoded all my projects in C++. 
