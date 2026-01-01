#------ Importing Libraries ---------
import os
from turtle import Turtle,Screen
import pandas as pd
from tkinter import simpledialog
import pygame
pygame.mixer.init()
import time
import random

# ------ Screen & Turtle Initialization ------
pen = Turtle()
screen = Screen()

# ------ Displaying Project Features in Console ------
print("\n----------------------------------------------------------\n\t\t~-*-~ Machine Features ~-*-~")
print(f"""----------------------------------------------------------
    1 -> Add multiple candidates with ease
    2 -> Unique symbol auto-assigned to each candidate
    3 -> Voter verification via voter-ID
    4 -> Prevents duplicate voting
    5 -> NOTA option available for neutral voting
    6 -> Admin-only access to end voting
    7 -> Secure result display with password
    8 -> Tie detection and winner announcement
----------------------------------------------------------""")
print("\n\t-~-~-: ELECTRONIC VOTING MACHINE :-~-~-\n")

# ------ Variables Used Throughout Program ------
candidate_list = []         # Stores candidate names
candidate_votes = []        # Stores vote count of each candidate
candidate_sign = []         # Stores unique symbol of each candidate

voter_id_list = []          # Tracks voters who already voted
voter_name_list = []        # Stores voter names
voter_sign = []             # Stores symbol voted by each voter

password = "EVM@1234"       # Admin password to stop voting

sign = ["!", "@", "#", "$", "%", "^", "&", "*", "?", "/", "-", "+", "<", ">", ":", ";", "|", "=", "}", "{"]
FONT = ("Arial",18,"bold")
ER_FONT = ("Fixedsys",18,"bold")
vote_successfully_voice = pygame.mixer.Sound("../other_files/beep.mp3")
condition = True
voting = True

# ------ Screen Configuration ------
screen.title("Electronic Voting Machine")
screen.setup(1000,700)
screen.bgpic("../other_files/bg_evm.gif")

# ------ Root window access (needed for password dialog) ------
root = screen._root

# ------ Additional Turtle for Messages ------
pen2 = Turtle()
pen.hideturtle()
pen2.hideturtle()
pen.penup()
pen.shapesize(3)
pen.setpos(0,300)
pen.color("red")

# ------ Project Title ------
pen.write("--- Welcome to Electronic Voting Machine ---",align="Center",font=("Times New Roman",18,"bold"))

# ------ Election Name Input ------
pen.setpos(0,250)
pen.color("blue")
election_name = screen.textinput("Election Name","Enter your Election name").title()
pen.write(f"{election_name} Election",align="Center",font=("Times New Roman",18,"bold"))
pen2.penup()
pen2.setpos(0,50)
pen2.color("red")

# ------ Load Voter Data from Excel ------
file = "C:/Users/91788/Downloads/voter_list.xlsx"
data = pd.read_excel(file)
voter_ids = [x.strip() for x in data["Epic_no"].to_list()]
voter_names = [x.strip() for x in data["Name Of Voters"].to_list()]

# ------ Candidate Registration Logic ------
while condition:
    time.sleep(1)
    pen2.clear()
    try:
        candidate_number = int(screen.textinput("Total Candidate",f"Q.How many candidates standing for {election_name} Election ?"))
        if candidate_number>20:
            pen2.clear()
            pen2.write("20 Candidate allows",font=FONT,align="center")

        elif candidate_number == 1:
            pen2.clear()
            pen2.write("Atleast enter two candidate name!",font=FONT,align="center")

        elif candidate_number == 0:
            pen2.clear()
            pen2.write("ERROR!,Machine stopped!",font=FONT,align="center")
            condition = False
            voting = False

            # Collect candidate names and assign unique symbols
        else:
            for x in range(candidate_number):
                candidate_name = screen.textinput("Enter Candidate name : ",f"Enter name of Candidate No.{x + 1} : ").lower().title()
                candidate_list.append(candidate_name)
                candidate_votes.append(0)
                sign_candy = random.choice(sign)
                candidate_sign.append(sign_candy)
                sign.remove(sign_candy)

            # Adding NOTA option
            candidate_list.append("NOTA")
            candidate_votes.append(0)
            candidate_sign.append("N")
            condition = False

    except ValueError:
            pen2.clear()
            pen2.write("Invalid input!!!,plz enter value in number!",font=FONT,align="center")
pen2.clear()

# ------ Display Candidate Table ------
Y = 150
screen.tracer(0)
pen.color("purple")
pen.setpos(-480,190)
pen.write("Sr.no",font=FONT)
pen.forward(100)
pen.write("Candidate Name",font=FONT)
pen.forward(300)
pen.write("Candidate Sign",font=FONT)

# Loop to display each candidate
for x in range(len(candidate_list)):
    pen.setpos(-460, Y)
    pen.color("blue")
    pen.write(f"{x+1}",font=FONT)
    pen.setpos(-340, Y)
    pen.color("green")
    pen.write(candidate_list[x],font=FONT)
    pen.setpos(0,Y)
    pen.color("black")
    pen.write(candidate_sign[x],font=FONT)
    Y-=40

screen.update()
time.sleep(1)
y_error = -140

# ------ Voting Logic ------
error = Turtle()
error.penup()
error.hideturtle()
error.setpos(0,y_error)
colors = ["red","blue","green","black","yellow","purple","orange",]

while voting:
    time.sleep(3)
    voter_id = screen.textinput("Candidate vote","\nEnter your voter-id to vote or (type '1234' to stop) : ").lower().capitalize()
    error.color(random.choice(colors))

    # Case 1: Duplicate vote
    if voter_id in voter_id_list:
        error.clear()
        error.write("You Already vote",align="center",font=ER_FONT)

    # Case 2: Stop voting & show result
    elif voter_id == "1234":
        password_ = simpledialog.askstring("Password to Show Result","Enter admin password : ",parent=root,show="*").upper()
        if password_ == password:
            voting = False
            screen.clear()

            # Display result table
            screen.tracer(0)
            pen.setpos(0, 300)
            pen.color("red")
            pen.write("--- Welcome to Electronic Voting Machine ---", align="Center",font=("Times New Roman", 18, "bold"))
            pen.setpos(0, 250)
            pen.color("blue")
            pen.write(f"{election_name} Election", align="Center", font=("Times New Roman", 18, "bold"))
            pen.setpos(-450,200)
            pen.color("purple")
            pen.write("Candidate Name",font=FONT)
            pen.forward(400)
            pen.write("Vote",font=FONT)
            winner = max(candidate_votes)
            y = 160
            for x in range(len(candidate_list)):
                pen.color("blue")
                pen.setpos(-420, y)
                pen.write(candidate_list[x], align="left", font=FONT)
                pen.color("green")
                pen.setpos(-25, y)
                pen.write(f"{candidate_votes[x]}/{sum(candidate_votes)}", align="center", font=FONT)
                y -= 40
            screen.update()
            winners = []
            for i in range(len(candidate_votes)):
                if candidate_votes[i] == winner:
                    winners.append(candidate_list[i])

            if len(winners) == 1:
                error.clear()
                error.color("red")
                error.write(f" Winner : {winners[0]}",align="center",font=ER_FONT)
            else:
                error.clear()
                error.write(f"Result has been Tied between ",align="center",font=ER_FONT)
                error.color("yellow")
                for w in winners:
                    y_error -= 40
                    error.setpos(0,y_error)
                    error.color("red")
                    error.write(f"{w}", align="center", font=ER_FONT)
        else:
            error.clear()
            error.write(f"Incorrect password!", align="center", font=ER_FONT)

    # Case 3: Invalid voter ID
    elif voter_id not in voter_ids:
        error.clear()
        error.write(f"Invalid Voter-ID", align="center", font=ER_FONT)

    # Case 4: Valid vote
    else:
        vote = screen.textinput("Register Your Vote","Enter your candidate sign to vote : ")
        if vote in candidate_sign:
            score = candidate_sign.index(vote)
            candidate_votes[score] += 1
            voter_id_list.append(voter_id)
            voter_sign.append(vote)
            index_name = voter_ids.index(voter_id)
            voter_name_list.append(voter_names[index_name])
            error.clear()
            vote_successfully_voice.play()
            error.write("Vote recorded Successful!", align="center", font=ER_FONT)

        else:
            error.clear()
            error.write("Invalid sign!", align="center", font=ER_FONT)
time.sleep(2)
screen.exitonclick()