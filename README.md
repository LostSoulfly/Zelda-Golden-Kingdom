# Zelda-Golden-Kingdom
Fan-made Zelda online role playing game made in VB6.

# Update 2023
Unfortunately the translation component I wrote no longer functions and will crash when encountering any untranslated text.
I will need to re-install VB6 and develop a new C# translation DLL and possibly use an old version of Visual Studio with .NET 4.0 framework to make it fully functional, but you can still play a lot of the game as it is now.
I may do this some day.. but today is not that day.

## Help
1. In the Cliente folder, install the Libraries.exe to get the VB6 required files.
2. Then run Starter.exe as Administrator to register the translation DLLs.
3. Then you should be able to start the Zelda client executable.
4. If you get a launcher required error, edit the config.ini to: `RequireLauncher= 0`
5. Now, start the Server executable in the server folder. (Server 0.72.exe)

### Original

I don't understand git. Sorry if it's a mess.

I found this sourcecode online on some forums, turned out to be a near-complete
online game built using a modified version of the Eclipse game engine made in VB6.

It was originally made in Spanish. I don't speak Spanish.

I managed to create a C# DLL that utilizes Bing, Google, and Yandex's online translation API/services.

This turned out to work great. It's a little more involved to setup and register the DLL, but nothing too rough.

Once it's registered, you can simply call the DLL's exported functions. See modTranslate.bas on server/client.

I also wrote a simple language database for storing the translations in a cache of sorts.

After doing that, I wrote a translation editor for reading and modifying that database.
It's not complete, but worked for my purposes.

This also includes the updater that I got off of the Eclipse forums and modified, as it was quick and easy.

I have a server running currently, you can use this source to connect to trollparty.org port 4000 for the normal server.

Or download the launcher here: http://trollparty.org/Zelda/Launcher.zip
