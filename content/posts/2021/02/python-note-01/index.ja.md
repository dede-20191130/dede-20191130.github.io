---
title: "[Python] 久しぶりにpythonスクリプトを作成した。[ファイル区分けスクリプト]"
author: dede-20191130
date: 2021-02-12T23:23:50+09:00
slug: python-note-01
draft: false
toc: true
featured: false
tags: ["python"]
categories: ["雑記"]
archives:
    - 2021
    - 2021-02
---

## About

ストレージから移したファイル区分けのために久しぶりにPythonスクリプトを組んだ。  
忘備録ついでに記したい。

Pythonは自由度が高いし  
IDEでの補完機能が弱い（動的型付けなので仕方ないが）ことを除けば、  
扱っていて楽しい言語だと思う。

## 環境

- Windows10
- Python 3.7.3

## 用途

10000個程度ある画像ファイルを名前順でソートして別のフォルダにコピーする。  
移す際にサブフォルダを作成し、数百個ずつくらいの組にして仕分けする。

各画像ファイルのファイル名が1.jpg~12000.jpg(.png or .mov)などで規則的だったので   
名前順でソートすることにした。

## コード

```py
# FileName:fileSorting.py
import glob
import os
import shutil
from win10toast import ToastNotifier


def Main():
    isSuccess = True

    try:
        mySrcPath = r"C:\tmp\20210211\src"
        myDestpath = r"C:\tmp\20210211\dst"
        movedFileList = []

        # 指定フォルダ直下のすべてのファイルをListに取得
        # ファイルの名前（6桁ゼロ埋め）でソート
        if os.path.isdir(mySrcPath):
            for file in sorted(
                    glob.glob(mySrcPath + '\\*', recursive=False),
                    key=lambda my_file: os.path.splitext(os.path.basename(my_file))[0].zfill(6)
            ):

                if os.path.isfile(file):
                    movedFileList.append(file)
        else:
            print('No Directory')
            exit(0)

        currentDirNum = 1
        # 200個ずつ区分け
        bundleNum = 200
        imgNum = len(movedFileList)
        stt = 1
        end = 0

        # currentDirNumの番号通りの名前のサブフォルダを作成し、bundleNum個ずつファイルをコピーして格納する
        while imgNum >= bundleNum:
            end += bundleNum
            newList = movedFileList[stt - 1:end]
            stt += bundleNum
            dst = myDestpath + "\\" + str(currentDirNum)
            os.makedirs(dst)
            currentDirNum += 1
            # # 指定出力先フォルダにコピー
            [shutil.copy2(os.path.abspath(file2), dst) for file2 in newList]
            imgNum -= bundleNum
        # 残余分のファイルをコピー
        newList = movedFileList[end:]
        if newList:
            dst = myDestpath + "\\" + str(currentDirNum)
            os.makedirs(dst)
            [shutil.copy2(os.path.abspath(file2), dst) for file2 in newList]

    except:
        isSuccess = False

    finally:
        toast = ToastNotifier()
        if isSuccess:
            toast.show_toast("処理完了",
                             "ファイルまとめが完了しました。",
                             icon_path=None,
                             duration=2,
                             threaded=True)
        else:
            toast.show_toast("処理エラー",
                             "正常に終了しませんでした。",
                             icon_path=None,
                             duration=2,
                             threaded=True)


if __name__ == '__main__':
    Main()

```