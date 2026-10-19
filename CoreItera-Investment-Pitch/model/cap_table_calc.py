S=1_000_000; R=100_000; cash=30.0; tj_cash=20.0; rk_cash=10.0; cap=0.20
def show(name, shares):
    t=sum(shares.values()); print(name, {k:f"{v:,.0f} ({v/t:.2%})" for k,v in shares.items()}, f"total {t:,.0f}")
# Option B: grant to 15% then cash to 20% combined
G=(0.15*S-R)/(1-0.15); S1=S+G
C=(cap*S1-(R+G))/(1-cap); price=cash*1e6/C
tj=C*tj_cash/cash; rk=C*rk_cash/cash
print("B grant",round(G),"cash shares",round(C),"price/share",round(price,1),"pre(after grant) Rp%.0fM"%(S1*price/1e6),"post Rp%.0fM"%((S1+C)*price/1e6))
show("B before", {"Others":S-R,"Riskia":R})
show("B after grant", {"Others":S-R,"Riskia":R+G})
show("B after cash", {"Others":S-R,"Riskia":R+G+rk,"TJ":tj})
# Option A: cash only
n=(cap-R/S)/(1-R/S); C=n*S/(1-n); price=cash*1e6/C
print("A new share of post %.2f%%"%(n*100),"pre Rp%.0fM post Rp%.0fM"%(S*price/1e6,(S+C)*price/1e6))
show("A after", {"Others":S-R,"Riskia":R+C*rk_cash/cash,"TJ":C*tj_cash/cash})
# runway
burn=6.6; print("runway months", cash/burn)
for seat in (250,350,500):
    print("seat Rp%dk -> paying seats to cover burn: %.1f"%(seat, burn*1000/seat))
