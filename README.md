# Zelda-Golden-Kingdom
Fan-made Zelda online role playing game made in VB6.


# Update 2025
I did a lot of work on updating the client back in late 2023, but ran into some issues. I don't remember what they were and as such I'm releasing this update as a testing version with no plans to continue.

I created a few additional tools for the translation including a program to read the source files and apply automatic translations rather than doing it on-the-fly. I also wrote a TranslationServer which communicates with the client/server programs to facilitate translation without the need of the GTranslate.DLL or the local translation databases. I've included those tools. I also did a quick update to the Launcher and added a web server to wrap up the package nicely. If the launcher doesn't connect after starting the web server, close and re-open it.

To run your own server and client, you should clone this repo and run the `Launcher.exe` from the Client folder. In the Launcher, click `Start WebSvr` and `Start GameSvr`. You can also run `Start.bat` in the root folder to start everything.

To play with others, you can update the `Client\data\UpdateConfig.ini` to point to your IP address.

You can also run the `TranslationServer.exe` and it will connect to the Client and Server, but I don't think there's any code implemented to use it any more, since I was using it to translate all the text manually when saving it from within the server.

![image](https://github.com/user-attachments/assets/e4fbf37d-c0ed-4dc6-a320-b37997942635)


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
