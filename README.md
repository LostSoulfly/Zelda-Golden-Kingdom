# Zelda-Golden-Kingdom
Fan-made Zelda online role playing game made in VB6.

I don't understand git. Sorry if it's a mess.

I found this sourcecode online on some forums, turned out to be a near-complete
online game built using a modified version of the Eclipse game engine made in VB6.

It was originally made in Spanish. I don't speak Spanish.

I managed to create a C# DLL that utilizes Bing, Google, and Yandex's online translation API/seervices.

This turned out to work great. It's a little more involved to setup and register the DLL, but nothing too rough.

Once it's registered, you can simply call the DLL's exported functions. See modTranslate.bas on server/client.

I also wrote a simple language database for storing the translations in a cache of sorts.

After doing that, I wrote a translation editor for reading and modifying that database.
It's not complete, but worked for my purposes.

This also includes the updater that I got off of the Eclipse forums and modified, as it was quick and easy.

I have a server running currently, you can use this source to connect to trollparty.org port 4000 for the normal server, or port 4001 for the 'Troll' server where everyone is an admin (Banning/kicking is disabled.)

Or download the launcher here: http://trollparty.org/Zelda/Launcher.zip
