# FitWallpaper
Windows light application for change wallpaper periodically

* No UI
* Wallpaper only applied via Fit mode
* Image is selected randomly in folder
* 3 Options for the empty space color (Black, White, __Dominant Color__)
* Choose change wallpaper period between 15 mins and 1 day inclusive
* __This program don't support multiple display__

# Requirements
* Windows XP SP1 or later with x64 (not include ARM64) version
* __A lots of picture files__ in single folder
* [Microsoft Visual Studio Community 2019](https://visualstudio.microsoft.com/downloads/)
* [OpenCV 3.4.14 or later - Windows x64 build](https://opencv.org/releases/)

# How to build
1. Open Visual Studio and create __empty windows desktop application__ solution named __FitWallpaper__.
2. Change build target to __x64 - Release__.
3. Add __FitWallpaper.cpp__.
4. Link __OpenCV header and lib__. You may google about this. It's easy.  
   But __don't use opencv_world[version]d.lib__ which for debug.
5. Build.
6. Copy __FitWallpaper.exe__ in FitWallpaper\x64\Release and  
   __opencv_world[version].dll__ in OpenCV\build\x64\vc15\bin to any location you like.  
   __Don't use opencv_world[version]d.dll__ which for debug.
7. Run __FitWallpaper.exe__.

# Usage
* Run program will register __run at startup__ automatically.
* If there are no config file, program create it and exit. You have to __edit it__.
* For uninstall, run __stop.cmd__.
* For reload config, run __stop.cmd__ then run __FitWallpaper.exe__ again.
* Whether program is running or not, you can change wallpaper by __drag single image file__ into __FitWallpaper.exe__.
* While program is running, you can change wallpaper immediately by run __FitWallpaper.exe__ again.

# Trivia
* __I know single source code is really bad__.  
  But I didn't think this program could be over 1000 lines... Sorry about this point.  
