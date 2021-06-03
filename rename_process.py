import glob
import os


def rename():
    try:
        FCE = r"Text File/FCE.txt"
        FCE_a = r"Text File/FCE.a.txt"
        NBA = r"Text File/NBA.txt"
        NBA_a = r"Text File/NBA.a.txt"
        LVC = r"Text File/LVC.txt"
        LVC_a = r"Text File/LVC.a.txt"
        LBU = r"Text File/LBU.txt"
        LBU_a = r"Text File/LBU.a.txt"
        NBC = r"Text File/NBC.txt"
        NBC_a = r"Text File/NBC.a.txt"
        MLG1 = r"Text File/MLG-01.txt"
        MLG1_a = r"Text File/MLG-01.a.txt"
        MLG2 = r"Text File/MLG-02.txt"
        MLG2_a = r"Text File/MLG-02.a.txt"
        MLG3 = r"Text File/MLG-03.txt"
        MLG3_a = r"Text File/MLG-03.a.txt"
        MLG6 = r"Text File/MLG-06.txt"
        MLG7 = r"Text File/MLG-07.txt"
        MLG7_a = r"Text File/MLG-07.a.txt"

        if os.path.exists(FCE):
            os.rename(FCE, r"Text File/A.FCE.txt")
        else:
            pass
        if os.path.exists(FCE_a):
            os.rename(FCE_a, r"Text File/B.FCE.a.txt")
        else:
            pass
        if os.path.exists(NBA):
            os.rename(NBA, r"Text File/C.NBA.txt")
        else:
            pass
        if os.path.exists(NBA_a):
            os.rename(NBA_a, r"Text File/D.NBA.a.txt")
        else:
            pass
        if os.path.exists(LVC):
            os.rename(LVC, r"Text File/E.LVC.txt")
        else:
            pass
        if os.path.exists(LVC_a):
            os.rename(LVC_a, r"Text File/F.LVC.a.txt")
        else:
            pass
        if os.path.exists(LBU):
            os.rename(LBU, r"Text File/G.LBU.txt")
        else:
            pass
        if os.path.exists(LBU_a):
            os.rename(LBU_a, r"Text File/H.LBU.a.txt")
        else:
            pass
        if os.path.exists(NBC):
            os.rename(NBC, r"Text File/I.NBC.txt")
        else:
            pass
        if os.path.exists(NBC_a):
            os.rename(NBC_a, r"Text File/J.NBC.a.txt")
        else:
            pass
        if os.path.exists(MLG1):
            os.rename(MLG1, r"Text File/K.MLG1.txt")
        else:
            pass
        if os.path.exists(MLG1_a):
            os.rename(MLG1_a, r"Text File/L.MLG1.a.txt")
        else:
            pass
        if os.path.exists(MLG2):
            os.rename(MLG2, r"Text File/M.MLG2.txt")
        else:
            pass
        if os.path.exists(MLG2_a):
            os.rename(MLG2_a, r"Text File/N.MLG2.a.txt")
        else:
            pass
        if os.path.exists(MLG3):
            os.rename(MLG3, r"Text File/O.MLG3.txt")
        else:
            pass
        if os.path.exists(MLG3_a):
            os.rename(MLG3_a, r"Text File/P.MLG3.a.txt")
        else:
            pass
        if os.path.exists(MLG6):
            os.rename(MLG6, r"Text File/Q.MLG6.txt")
        else:
            pass
        if os.path.exists(MLG7):
            os.rename(MLG7, r"Text File/R.MLG7.txt")
        else:
            pass
        if os.path.exists(MLG7_a):
            os.rename(MLG7_a, r"Text File/S.MLG7.a.txt")
        else:
            pass
    except FileExistsError:
        for new_txt in glob.glob("Text File\\*.txt"):
            os.remove(new_txt)
