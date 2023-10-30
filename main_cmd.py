from ortools.sat.python import cp_model
import icecream as ic
import calendar
import os
import xlsxwriter


MONTH = 11
YEAR = 2023
CUSTUM_DAYS = [21]


# Define the workers
workers = [
    {
        "name": "w1",
        "max_shifts": 3,
        "min_shifts": 0,
        "preferences": [1, 2, 3, 4, 5, 6, 7, 8, 9],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
    {
        "name": "w2",
        "max_shifts": 3,
        "min_shifts": 0,
        "preferences": [
            1,
            2,
            3,
            4,
            5,
            6,
            7,
            8,
            9,
            10,
            11,
            12,
            13,
            14,
            15,
            16,
            17,
            18,
            19,
        ],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
    {
        "name": "w3",
        "max_shifts": 3,
        "min_shifts": 0,
        "preferences": [],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
    {
        "name": "w4",
        "max_shifts": 3,
        "min_shifts": 0,
        "preferences": [],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
    {
        "name": "w5",
        "max_shifts": 4,
        "min_shifts": 0,
        "preferences": [],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
    {
        "name": "w6",
        "max_shifts": 4,
        "min_shifts": 0,
        "preferences": [],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
    {
        "name": "w7",
        "max_shifts": 4,
        "min_shifts": 0,
        "preferences": [],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
    {
        "name": "w8",
        "max_shifts": 4,
        "min_shifts": 0,
        "preferences": [],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
    {
        "name": "w9",
        "max_shifts": 4,
        "min_shifts": 0,
        "preferences": [],
        "total_shifts": 0,
        "shifts_on_holiday": 0,
    },
]


def find_holidays(year, month, custom_days):
    # Get the calendar for the specified month and year
    cal = calendar.monthcalendar(year, month)

    # Define the day numbers for Sunday and Saturday
    sunday = calendar.SUNDAY
    saturday = calendar.SATURDAY

    # Use list comprehension to generate the list of days
    days = [day for week in cal for day in (week[sunday], week[saturday]) if day != 0]

    # Add custom days to the list if provided
    if custom_days:
        days.extend(custom_days)

    return sorted(set(days))


holidays = find_holidays(YEAR, MONTH, CUSTUM_DAYS)

# based on the month and year, get the number of days in the month
num_days = (
    calendar.monthrange(YEAR, MONTH)[1] + 1
)  # +1 because the range is 0-30, not 1-31


def assign_workers(workers, days):
    num_shifts = 1
    num_days = days
    all_days = range(num_days)

    # Create the model
    model = cp_model.CpModel()

    # Create the variables
    shifts = {
        (worker["name"], day): model.NewBoolVar(f"{worker['name']}_{day}")
        for worker in workers
        for day in range(1, num_days)
    }

    # Create the constraints
    for worker in workers:
        # Each worker must work a maximum of x shifts
        model.Add(
            sum(shifts[(worker["name"], day)] for day in range(1, num_days))
            <= worker["max_shifts"]
        )

        # Each worker must work a minimum of x shifts
        model.Add(
            sum(shifts[(worker["name"], day)] for day in range(1, num_days))
            >= worker["min_shifts"]
        )

        # Each worker must not work more than 1 shift in a row
        for day in range(1, num_days - 1):
            model.Add(
                shifts[(worker["name"], day)] + shifts[(worker["name"], day + 1)] <= 1
            )

        # Each worker must work at most  one time per month on holidays
        model.Add(
            sum(
                shifts[(worker["name"], day)]
                for day in holidays
                if day not in worker["preferences"]
            )
            <= 1
        )

    # Each day must have exactly 1 worker
    for day in range(1, num_days):
        model.AddExactlyOne([shifts[(worker["name"], day)] for worker in workers])

    # Try to distribute the shifts evenly, so that each worker works
    # min_shifts_per_worker shifts. If this is not possible, because the total
    # number of shifts is not divisible by the number of nurses, some workers will
    # be assigned one more shift.

    # min_shifts_per_worker = (num_shifts * num_days) // num_workers

    # Each worker must not work on their preferences
    model.Maximize(
        sum(
            shifts[(worker["name"], day)]
            for worker in workers
            for day in range(1, num_days)
            if day not in worker["preferences"]
        )
    )

    # Create the solver
    solver = cp_model.CpSolver()

    # Solve the model
    status = solver.Solve(model)

    # Check if the model is feasible
    if status == cp_model.FEASIBLE or status == cp_model.OPTIMAL:
        # Create the results
        results = {}
        for day in range(1, num_days):
            for worker in workers:
                if solver.Value(shifts[(worker["name"], day)]) == 1:
                    results[day] = worker["name"]
                    break

        # Count the shifts
        for worker in workers:
            worker["total_shifts"] = sum(
                solver.Value(shifts[(worker["name"], day)])
                for day in range(1, num_days)
            )
            worker["shifts_on_holiday"] = sum(
                solver.Value(shifts[(worker["name"], day)]) for day in holidays
            )

        return results
    else:
        return None


def export_to_excel(workers, shifts, filename, holidays):
    # Check if the file exists and remove it if it does
    if os.path.exists(filename):
        os.remove(filename)
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Define a bold format to use for headers.
    bold = workbook.add_format({"bold": True})

    # Define a format for prefrenced days (light red fill)
    prefrenced_format = workbook.add_format(
        {"bold": True, "bg_color": "#ffcccc", "border": 1}
    )

    # Define a format for holidays (darker gray fill) and bold
    holiday_format = workbook.add_format(
        {"bold": True, "bg_color": "#d3d3d3", "border": 1}
    )

    # Set the column width for the Employee Names and Total Shift Count and Holiday Count columns
    worksheet.set_column("A:A", 20)
    worksheet.set_column("B:B", 10)
    worksheet.set_column("C:C", 10)

    # Set the column width for the days
    worksheet.set_column("D:AI", 3)

    # Sort the workers alphabetically
    workers.sort(key=lambda x: x["name"])

    col = 0
    row = 1

    # Write the headers for Names and Shift Count
    worksheet.write("A1", "ΟΝΟΜΑ", bold)
    worksheet.write("B1", "ΣΥΝΟΛΟ", bold)
    worksheet.write("C1", "ΑΡΓΙΕΣ", bold)

    # Write employee data
    for worker in workers:
        worksheet.write(row, col, str(worker["name"]))
        worksheet.write(row, col + 1, worker["total_shifts"])
        worksheet.write(row, col + 2, worker["shifts_on_holiday"])
        row += 1

    col = 4
    row = 1

    # Write the days as headers
    for day in range(num_days - 1):
        day_label = day + 1
        header_format = bold if day_label not in holidays else holiday_format
        worksheet.write(0, col + day, day_label, header_format)

    # Write the schedule for each worker
    for worker in workers:
        worker_schedule = [
            "X" if shifts[day] == worker["name"] else "" for day in range(1, num_days)
        ]
        shift_format = workbook.add_format({"border": 1, "align": "center"})
        worksheet.write_row(row, col, worker_schedule, shift_format)
        row += 1

    # make the cells with the prefrenced days light red
    col = 4
    row = 1
    for worker in workers:
        for day in worker["preferences"]:
            worksheet.write(row, col + day - 1, "", prefrenced_format)
        row += 1
    # write the shifts one under the other in the format Day: Employee
    row = len(workers) + 2
    col = 0
    worksheet.write(len(workers) + 2, col, "ΑΝΑΛΥΣΗ ΑΝΑ ΜΕΡΑ", bold)

    for row, day in enumerate(range(1, num_days), len(workers) + 3):
        header_format = bold if day not in holidays else holiday_format
        worksheet.write(row, col, f"{day}: {shifts[day]}", header_format)

    # Close the workbook.
    workbook.close()


def results():
    # Assign the workers
    assigned_workers = assign_workers(workers, num_days)

    # Print the results
    if assigned_workers:
        for day in range(1, num_days):
            print(f"Day {day}: {assigned_workers[day]}")

        sum = 0
        for worker in workers:
            print(
                f"{worker['name']} >> {worker['total_shifts']} / {worker['max_shifts']} and {worker['shifts_on_holiday']} on holidays"
            )
            sum += worker["total_shifts"]
        print(f"Total shifts: {sum}")

        # total holidays
        sum = 0
        for worker in workers:
            sum += worker["shifts_on_holiday"]
        print(f"Total holidays: {sum}")

        # Export the results to an Excel file
        export_to_excel(workers, assigned_workers, "schedule.xlsx", holidays)

    else:
        print("No solution found.")


if __name__ == "__main__":
    results()
