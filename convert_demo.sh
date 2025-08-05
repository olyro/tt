#!/bin/bash

NODE_OPTIONS="--max-old-space-size=10000" svg-term --cast=$1 --out demo.svg --window --term iterm2 --profile "./iTerm2 Solarized Light.itermcolors"
