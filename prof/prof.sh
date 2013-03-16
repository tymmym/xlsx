#!/bin/bash

dist/build/prof/prof +RTS -K200M -p -hc -sprof.summ
hp2ps -c prof.hp
