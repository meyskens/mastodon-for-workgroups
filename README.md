# Mastodon 3.11 For Workgroups

This project is a Mastodon client written in Visual Basic 6. It works on Windows 95 and higher (Windows 10/11 untested). 

This project is in very early stages!! Use at own risk. Contributions very welcome! (If you can not use Git email me for an address to mail your patched on floppy disk to).

What works:

- Sending a toot
- Loading 20 posts of your home timeline

What is planned soon:

- Image support
- Avatars
- Boosts and likes
- Replies
- Refreshing toots without crashing

What I can use help on:

- Tabs for different timelines
- Better errorhandling
- More things...


## Shut up and take my floppy

If you are in no mood to install VB6 I understand. There is an installer [under releases](https://github.com/meyskens/mastodon-for-workgroups/releases/download/alpha-1/mfw-windows.9x.zip)

## What do I need?

You need a HTTPS to HTTP proxy, one that preferably also converts UTF-8 to Windows encoding. I use [WebOne](https://github.com/atauenis/webone) for this.
As this project sens your persona token always host the proxy yourself.

Once you set the proxy as your system proxy in Internet Explorer it will work. Press the "refresh" button to log in to your mastodon instance.

## Demos

- [Reading posts](https://blahaj.social/@maartje/109372878061833398)
- [Posting a toot](https://blahaj.social/@maartje/109376527177239374)
