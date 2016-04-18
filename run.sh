#!/bin/bash 
SRC_PATH=$PWD
docker run -ti --rm -u admin -w /src -v $PWD:/src -e DISPLAY=$DISPLAY -v /tmp/.X11-unix:/tmp/.X11-unix jimlin95/python3-env
