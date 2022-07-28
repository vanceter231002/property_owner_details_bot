import time
def loading_animation(name):
    animation=r"|/-\\"
    idx=0
    while(True):
        i=idx%len(name)
        temp=name[:i]+name[i].upper()+name[i+1:]
        print(f"{temp}{animation[idx%5]}",end="\r")
        idx+=1
        time.sleep(0.1)
    