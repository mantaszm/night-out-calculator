import customtkinter as ctk
from tkinter import filedialog
import json
from openpyxl import Workbook

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Night Out Calculator")
        self.geometry("1350x850")

        self.people = []
        self.items = []
        self.bank_person = None

        self.setup_screen()

    # ---------------- SETUP ----------------
    def setup_screen(self):
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            self.frame,
            text="Enter Names (one per line)",
            font=ctk.CTkFont(size=22, weight="bold")
        ).pack(pady=20)

        self.text = ctk.CTkTextbox(self.frame, width=320, height=300)
        self.text.pack(pady=10)

        ctk.CTkButton(self.frame, text="Create", command=self.create_from_input).pack(pady=5)
        ctk.CTkButton(self.frame, text="Load File", command=self.load_file).pack(pady=5)

    # ---------------- CREATE ----------------
    def create_from_input(self):
        raw = self.text.get("1.0", "end").strip().split("\n")
        self.people = [p.strip() for p in raw if p.strip()]

        if not self.people:
            return

        self.bank_person = self.people[0]

        self.frame.destroy()
        self.build_main()

    # ---------------- LOAD ----------------
    def load_file(self):
        file = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not file:
            return

        with open(file, "r") as f:
            data = json.load(f)

        self.people = data["people"]
        self.items = data["items"]
        self.bank_person = data.get("bank", self.people[0])

        if hasattr(self, "frame"):
            self.frame.destroy()

        self.build_main()

    # ---------------- MAIN UI ----------------
    def build_main(self):
        self.main = ctk.CTkFrame(self)
        self.main.pack(fill="both", expand=True)

        top = ctk.CTkFrame(self.main)
        top.pack(fill="x", pady=10)

        ctk.CTkButton(top, text="Add Item", command=self.add_item).pack(side="left", padx=5)
        ctk.CTkButton(top, text="Save", command=self.save_file).pack(side="left", padx=5)
        ctk.CTkButton(top, text="Export Excel", command=self.export_excel).pack(side="left", padx=5)

        ctk.CTkLabel(top, text="Bank:").pack(side="left", padx=10)

        self.bank_var = ctk.StringVar(value=self.bank_person)
        ctk.CTkOptionMenu(
            top,
            variable=self.bank_var,
            values=self.people,
            command=self.set_bank
        ).pack(side="left")

        self.table = ctk.CTkScrollableFrame(self.main)
        self.table.pack(fill="both", expand=True, padx=10, pady=10)

        self.result_box = ctk.CTkFrame(self.main)
        self.result_box.pack(fill="x", padx=10, pady=10)

        self.render()

    # ---------------- BANK ----------------
    def set_bank(self, value):
        self.bank_person = value
        self.render()

    # ---------------- ADD ITEM ----------------
    def add_item(self):
        win = ctk.CTkToplevel(self)
        win.geometry("420x320")
        win.attributes("-topmost", True)
        win.grab_set()

        ctk.CTkLabel(win, text="Add Item", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)

        name = ctk.CTkEntry(win, placeholder_text="Item name")
        name.pack(pady=5)

        buyer = ctk.StringVar(value=self.people[0])
        ctk.CTkOptionMenu(win, variable=buyer, values=self.people).pack(pady=5)

        price = ctk.CTkEntry(win, placeholder_text="Price (€)")
        price.pack(pady=5)

        def save():
            try:
                p = float(price.get().replace(",", "."))
            except:
                return

            self.items.append({
                "name": name.get(),
                "buyer": buyer.get(),
                "price": p,
                "checks": {x: False for x in self.people}
            })

            win.destroy()
            self.render()

        ctk.CTkButton(win, text="Save", command=save).pack(pady=10)

    # ---------------- CALC ----------------
    def calculate(self):
        balances = {p: 0 for p in self.people}
        cost = {p: 0 for p in self.people}

        for item in self.items:
            involved = [p for p in self.people if item["checks"].get(p, False)]
            if not involved:
                continue

            share = item["price"] / len(involved)

            for p in involved:
                cost[p] += share
                balances[p] -= share

            balances[item["buyer"]] += item["price"]

        return balances, cost

    # ---------------- RESULTS ----------------
    def show_results(self):
        for w in self.result_box.winfo_children():
            w.destroy()

        balances, cost = self.calculate()

        ctk.CTkLabel(
            self.result_box,
            text="Night Out Cost Breakdown",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w")

        for p in self.people:
            if p == self.bank_person:
                continue

            ctk.CTkLabel(
                self.result_box,
                text=f"{p} | Cost: €{cost[p]:.2f} | Net: €{balances[p]:.2f}",
                anchor="w"
            ).pack(anchor="w")

    # ---------------- SAVE ----------------
    def save_file(self):
        file = filedialog.asksaveasfilename(defaultextension=".json")
        if not file:
            return

        with open(file, "w") as f:
            json.dump({
                "people": self.people,
                "items": self.items,
                "bank": self.bank_person
            }, f)

    # ---------------- EXCEL EXPORT (FIXED + SAFE) ----------------
    def export_excel(self):
        file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")]
        )

        if not file:
            return

        try:
            wb = Workbook()
            wb.remove(wb.active)

            balances, cost = self.calculate()

            for person in self.people:
                ws = wb.create_sheet(title=person[:30])

                ws.append(["Prekė", "Ar pirkai", "Tavo dalis (€)", "Tavo sumokėta (€)", "Balansas (€)"])

                total_paid = 0
                total_cost = 0

                for item in self.items:
                    involved = [p for p in self.people if item["checks"].get(p, False)]
                    if not involved:
                        continue

                    share = item["price"] / len(involved)

                    paid_val = item["price"] if item["buyer"] == person else 0
                    cost_share = share if person in involved else 0
                    net = paid_val - cost_share

                    total_paid += paid_val
                    total_cost += cost_share

                    ws.append([
                        item["name"],
                        "Taip" if paid_val else "Ne",
                        round(cost_share, 2),
                        round(paid_val, 2),
                        round(net, 2)
                    ])

                ws.append([])
                ws.append(["IŠ VISO", "", round(total_cost, 2), round(total_paid, 2), round(total_paid - total_cost, 2)])

                ws.append([])
                ws.append(["PAAIŠKINIMAS"])
                ws.append([f"Tu sumokėjai: {round(total_paid, 2)} €"])
                ws.append([f"Tavo reali vakaro kaina: {round(total_cost, 2)} €"])

                diff = total_paid - total_cost

                if diff > 0:
                    ws.append([f"Tau turi būti grąžinta: {round(diff, 2)} €"])
                else:
                    ws.append([f"Tu turi sumokėti: {round(abs(diff), 2)} €"])

            wb.save(file)
            print("Excel file saved:", file)

        except Exception as e:
            print("❌ Excel export failed:")
            print(e)

    # ---------------- RENDER ----------------
    def render(self):
        for w in self.table.winfo_children():
            w.destroy()

        headers = ["Item", "Buyer", "Price"] + self.people

        for c, h in enumerate(headers):
            ctk.CTkLabel(
                self.table,
                text=h,
                font=ctk.CTkFont(weight="bold")
            ).grid(row=0, column=c, padx=10, pady=10)

        for r, item in enumerate(self.items, start=1):
            ctk.CTkLabel(self.table, text=item["name"]).grid(row=r, column=0)
            ctk.CTkLabel(self.table, text=item["buyer"]).grid(row=r, column=1)
            ctk.CTkLabel(self.table, text=f"€{item['price']:.2f}").grid(row=r, column=2)

            for c, p in enumerate(self.people, start=3):
                var = ctk.BooleanVar(value=item["checks"].get(p, False))

                box = ctk.CTkFrame(self.table, fg_color="transparent")
                box.grid(row=r, column=c, padx=10, pady=5)

                chk = ctk.CTkCheckBox(
                    box,
                    text="",
                    variable=var,
                    command=lambda it=item, pp=p, v=var: it["checks"].update({pp: v.get()})
                )
                chk.pack()

        self.show_results()


if __name__ == "__main__":
    app = App()
    app.mainloop()