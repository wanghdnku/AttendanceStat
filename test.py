#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import platform

path_separator = '/' if platform.system() == 'Darwin' else '\\'

if __name__ == '__main__':
    print(path_separator)
