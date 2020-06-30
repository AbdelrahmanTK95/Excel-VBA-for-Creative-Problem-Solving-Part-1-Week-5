# Excel-VBA-for-Creative-Problem-Solving-Part-1-Week-5
Problem Statement
You run a one-man or one-woman shipping company using an 18-wheeler (semi) and your profit depends upon several factors. 
Profit is equal to revenue minus costs. Revenues are generally fixed (and dependent on a lot of economic factors) but costs can be minimized to increase profit.
One of the biggest costs with the shipping industry is that of fuel. Obviously, the more fuel efficient the vehicle the lower fuel costs will be. 
Semis are most efficient at 55 mph but as the speed increases the drag forces increase exponentially, which decreases fuel efficiency. 
On the flip side of the coin, however, is the fact that if you drive slower, then you won’t get as many shipments delivered in a certain amount of time. 
As you many have guessed, there is some optimization that can occur in this industry, and driving the right speed can have a dramatic effect on profits.
Total profit (P) is equal to total revenue (TR) minus operating costs (C):

P=TR-C

Total revenue ($/hr) is equal to the revenue per delivery (r, $/delivery) divided by the miles required per delivery 
(d, miles/delivery, we are neglecting the return trip in this analysis) multiplied by the driving speed (s, in miles/hour):

TR=rs/d

Operating costs (C, we are neglecting taxes, maintenance/repair costs, and other expenses in this simplistic approach) are equal to the cost of fuel:

C=fs/e

Here, f is the fuel cost ($/gallon) and e is the fuel efficiency of the truck (in miles/gallon). Note that both TR and C have units of $/hr.
Fuel cost (f) is generally fixed (you have no direct control over that although it does fluctuate over time, as you all know). 
The main variable that plays a role in profit is fuel efficiency. Fuel efficiency (e) is highly dependent on speed (s) as well as weight of the truck (w, in pounds):

e = (6.12 - w / 60000) * 0.92 ^ ((s - 55) / 5)

Your typical delivery route is 1000 miles, you charge $1500 per delivery, and gas costs $2.50/gallon.

Create a VBA function called truck(r, d, f, w, c) with arguments for revenue per delivery (r), distance required for delivery (d), fuel cost (f), truck weight (w), and output option (c). 
If c = TRUE then the function should output speed (s) at which the maximum profit occurs but if c = FALSE then the function will output the actual maximum profit (P).

Your function should use 20 rounds of the Golden Search method to find the speed that maximizes profit for the parameters that are input into the function. 
Speed is anticipated to be between 30 and 100 mph (but hopefully not actually that slow or fast!).

Some helpful hints:
The independent variable that you are working with is speed. The dependent variable is total profit (P). 
The question is asking for you to output speed (s) if the last argument in the function is TRUE; otherwise, if the last argument is FALSE, then the function should output the maximum profit.
Remember that the Golden Search method finds the minimum of a function. Our function of interest is profit, which we are trying to maximize. 
To maximize a function using Golden Search, we must find the minimum of the negative of the function we are trying to find the maximum of. 


