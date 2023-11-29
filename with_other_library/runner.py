from constraint import Problem
from ortools.sat.python import cp_model
import icecream as ic
import calendar


# Define variables (nurses and shifts)
nurses = [
    "Nurse1",
    "Nurse2",
    "Nurse3",
    "Nurse4",
    "Nurse5",
    "Nurse6",
    "Nurse7",
    "Nurse8",
    "Nurse9",
]
shifts = [
    "Day01",
    "Day02",
    "Day03",
    "Day04",
    "Day05",
    "Day06",
    "Day07",
    "Day08",
    "Day09",
    "Day10",
    "Day11",
    "Day12",
    "Day13",
    "Day14",
    "Day15",
    "Day16",
    "Day17",
    "Day18",
    "Day19",
    "Day20",
    "Day21",
    "Day22",
    "Day23",
    "Day24",
    "Day25",
    "Day26",
    "Day27",
    "Day28",
    "Day29",
    "Day30",
]

MONTH = 11
YEAR = 2023
CUSTUM_DAYS = [21]


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


def main():
    # Data.
    num_nurses = 9
    num_days = 30
    all_nurses = range(num_nurses)
    all_days = range(num_days)

    # Creates the model.
    model = cp_model.CpModel()

    # Creates shift variables.
    # shifts[(n, d)]: nurse 'n' works  on day 'd'.
    shifts = {}
    for n in all_nurses:
        for d in all_days:
            shifts[(n, d)] = model.NewBoolVar(f"shift_n{n}_d{d}")

    # Each day is assigned to exactly one nurse in the schedule period.
    for d in all_days:
        model.AddExactlyOne(shifts[(n, d)] for n in all_nurses)

    # Each nurse works at most one shift per day.
    for n in all_nurses:
        for d in all_days:
            model.AddAtMostOne(shifts[(n, d)])

    # Each nurse does not works two days in a row.
    for n in all_nurses:
        for d in range(1, num_days):
            model.AddBoolOr(
                [
                    shifts[(n, d - 1)].Not(),
                    shifts[(n, d)].Not(),
                ]
            )

    # Try to distribute the shifts evenly, so that each nurse works
    # min_shifts_per_nurse shifts. If this is not possible, because the total
    # number of shifts is not divisible by the number of nurses, some nurses will
    # be assigned one more shift.
    min_shifts_per_nurse = (num_days) // num_nurses
    if num_days % num_nurses == 0:
        max_shifts_per_nurse = min_shifts_per_nurse
    else:
        max_shifts_per_nurse = min_shifts_per_nurse + 1
    for n in all_nurses:
        shifts_worked = []
        for d in all_days:
            shifts_worked.append(shifts[(n, d)])
        model.Add(min_shifts_per_nurse <= sum(shifts_worked))
        model.Add(sum(shifts_worked) <= max_shifts_per_nurse)

    # Creates the solver and solve.
    solver = cp_model.CpSolver()
    solver.parameters.linearization_level = 0
    # Enumerate all solutions.
    solver.parameters.enumerate_all_solutions = True

    class NursesPartialSolutionPrinter(cp_model.CpSolverSolutionCallback):
        """Print intermediate solutions."""

        def __init__(self, shifts, num_nurses, num_days, limit):
            cp_model.CpSolverSolutionCallback.__init__(self)
            self._shifts = shifts
            self._num_nurses = num_nurses
            self._num_days = num_days
            self._solution_count = 0
            self._solution_limit = limit

        def on_solution_callback(self):
            self._solution_count += 1
            print(f"Solution {self._solution_count}\n")
            self.print_solution()

            if self._solution_count >= self._solution_limit:
                print(f"Stop search after {self._solution_limit} solutions")
                self.StopSearch()

        def solution_count(self):
            return self._solution_count

        def print_solution(self):
            """
            print the solution in the following format:

            | nurse |Total| 01 | 02 | 03 | 04 | 05 | ... | 28 | 29 | 30 |
            --------------------------------------------------------------
            | Nurse1|     |  1 |  0 |  0 |  0 |  0 | ... |  0 |  0 |  0 |
            | Nurse2|     |  1 |  0 |  0 |  0 | ... |  0 |  0 |  0 |  0 |
            | Nurse3|     |  0 |  0 |  0 |  0 | ... |  0 |  0 |  0 |  0 |

            where 1 means that the nurse works on that day.
            Total is the total number of shifts worked by the nurse per month.

            """

            # Print the header
            print("|nurse |Total|", end="")
            for d in all_days:
                print(f"{d+1:02d}|", end="")
            print()
            print("-" * (num_days * 3 + 20))

            # find the total number of shifts worked by each nurse
            nurses_total = {}

            # Print the rows
            for n in all_nurses:
                nurses_total[n] = 0
                for d in all_days:
                    nurses_total[n] += self.Value(self._shifts[(n, d)])
                print(f"|{nurses[n]:5}|{nurses_total[n]:5}|", end="")
                for d in all_days:
                    print(f"{self.Value(self._shifts[(n, d)]):02d}|", end="")
                print()

    # Display the first five solutions.
    solution_limit = 1
    solution_printer = NursesPartialSolutionPrinter(
        shifts, num_nurses, num_days, solution_limit
    )

    solver.Solve(model, solution_printer)

    # Statistics.
    print("\nStatistics")
    print(f"  - conflicts      : {solver.NumConflicts()}")
    print(f"  - branches       : {solver.NumBranches()}")
    print(f"  - wall time      : {solver.WallTime()} s")
    print(f"  - solutions found: {solution_printer.solution_count()}")


if __name__ == "__main__":
    main()
