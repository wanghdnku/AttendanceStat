#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import platform

path_separator = '/' if platform.system() == 'Darwin' else '\\'

print(path_separator)
