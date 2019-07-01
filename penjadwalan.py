import xlrd, random, xlsxwriter, os

# EXCEL INPUT FILE LOCATION
input_loc = ('input_data.xlsx')

# OPEN WORKBOOK
workbook = xlrd.open_workbook(input_loc)

# UJIAN
subject_sheet = workbook.sheet_by_index(0)
subject_count = subject_sheet.nrows - 1

# TIMESLOT
time_sheet = workbook.sheet_by_index(1)
time_count = time_sheet.nrows - 1

# RUANGAN
room_sheet = workbook.sheet_by_index(2)
room_count = room_sheet.nrows - 1


def init_population(n_individu):
    population = list()

    for n in range(n_individu):
        individual = init_individual()
        individual.insert(0, 0)
        population.append(list(individual))

    return population


def init_individual():
    chromosomes = list()

    for i in range(1, subject_count + 1):
        subject = list()
        subject.append(subject_sheet.cell_value(i, 0))
        subject.append(int(subject_sheet.cell_value(i, 2)))
        subject.append(random.randint(1, time_count))
        subject.append(random.randint(1, room_count))
        chromosomes.append(subject)

    return chromosomes


def calc_fitness(origin_pop):
    updated_population = list()

    for individual in origin_pop:

        unfit = False
        fitness = 0

        i = 1
        while i < len(individual):
            chr_a = list(individual[i])

            # CEK APAKAH KAPASITAS RUANGAN MEMADAI
            if chr_a[1] > room_sheet.cell_value(chr_a[3], 1):
                unfit = True
                fitness -= 1
            else:
                fitness += 1

            j = i + 1
            while j < len(individual):
                chr_b = list(individual[j])

                # CEK APAKAH ADA DI TIMESLOT YANG SAMA
                if chr_a[2] == chr_b[2]:
                    # CEK APAKAH ADA DI RUANG YANG SAMA
                    if chr_a[3] == chr_b[3]:
                        unfit = True
                        fitness -= 1

                j += 1

            i += 1

        if not unfit:
            fitness = 10000

        individual[0] = fitness

        updated_population.append(individual)

    return updated_population


def selection(population):
    parents = list()
    if population[0][0] > population[1][0]:
        parent_a = list(population[0])
        parent_b = list(population[1])
    else:
        parent_a = list(population[1])
        parent_b = list(population[0])

    highest = parent_a[0]
    higher = parent_b[0]

    if len(population) > 2:
        i = 2
        while i < len(population):
            if population[i][0] > highest:
                parent_b = list(parent_a)
                parent_a = list(population[i])
                higher = highest
                highest = population[i][0]
            elif population[i][0] > higher:
                parent_b = list(population[i])
                higher = population[i][0]
            i += 1

    parents.append(parent_a)
    parents.append(parent_b)

    return parents


def crossover(parents):
    parent_a = list(parents[0])
    parent_b = list(parents[1])

    offsprings = list()

    c_point = random.randint(1, len(parent_a) - 1)

    while c_point < len(parent_a):
        parent_a[c_point], parent_b[c_point] = parent_b[c_point], parent_a[c_point]

        c_point += 1

    parent_a[0] = parent_b[0] = 0
    offsprings.append(parent_a)
    offsprings.append(parent_b)

    return random.choice(offsprings)


def rand_time(individual, rapo):
    individual = list(individual)
    rand_chr = list(individual[rapo])
    rand_chr[2] = random.randint(1, time_count)

    individual[rapo] = rand_chr
    return individual


def rand_room(individual, rapo):
    individual = list(individual)
    rand_chr = list(individual[rapo])
    rand_chr[3] = random.randint(1, room_count)

    individual[rapo] = rand_chr
    return individual


def mutate(obj, mut_rate, n_pop):
    mutated_pop = list()
    mutated_pop.append(obj)

    mut_rate = int(mut_rate * (len(obj) - 1))

    n = 1

    while n < n_pop:

        m = 0
        new_obj = list(obj)

        while m < mut_rate:
            rand_point = random.randint(1, len(new_obj) - 1)

            new_obj = list(rand_time(new_obj, rand_point))
            new_obj = list(rand_room(new_obj, rand_point))

            m += 1

        n += 1
        mutated_pop.append(new_obj)

    return mutated_pop


def winner_exists(population):
    for ind in population:
        if ind[0] == 10000:
            return True


def get_winner(population):
    for ind in population:
        if ind[0] == 10000:
            return ind


if __name__ == '__main__':

    n_population = 5
    mutate_rate = 0.5
    winner = list()
    found = False

    pop = list(init_population(n_population))

    for ind in pop:
        print(ind)

    pop = list(calc_fitness(pop))

    print("POPULASI DENGAN FITNESS")
    for ind in pop:
        print(ind)

    if winner_exists(pop):
        found = True
        winner = get_winner(pop)

    g = 1
    while not found:
        print("generasi", g)
        parents = list(selection(pop))

        print("parents =")
        for parent in parents:
            print(parent)

        offspring = list(crossover(parents))

        print("anakan =\n", offspring)

        pop = list(mutate(offspring, mutate_rate, n_population))

        pop = list(calc_fitness(pop))

        print("hasil mutasi")
        for ind in pop:
            print(ind)

        if winner_exists(pop):
            found = True
            winner = get_winner(pop)

        g += 1

    print("Pemenang =", winner)

    workbook = xlsxwriter.Workbook(os.path.join(os.path.dirname(os.path.abspath(__file__)), "hasil-jadwal-ujian.xlsx"))
    worksheet = workbook.add_worksheet('Jadwal Ujian')
    bold = workbook.add_format({'bold': True})
    worksheet.write(0, 0, "JADWAL UJIAN AKHIR SEMESTER TEKNIK INFORMATIKA UNSOED 2019", bold)
    generasi = "Generasi ke-" + str(g)
    worksheet.write(2, 0, generasi)

    row = 5
    worksheet.write(row - 1, 0, "KODE", bold)
    worksheet.write(row - 1, 1, "PESERTA", bold)
    worksheet.write(row - 1, 2, "WAKTU", bold)
    worksheet.write(row - 1, 3, "RUANGAN", bold)

    i = 1
    while i < len(winner):
        worksheet.write(row, 0, winner[i][0])
        worksheet.write(row, 1, winner[i][1])
        time = time_sheet.cell_value(winner[i][2], 0) + ", " + time_sheet.cell_value(winner[i][2], 1) + " - " + time_sheet.cell_value(winner[i][2], 2)
        worksheet.write(row, 2, time)
        room = room_sheet.cell_value(winner[i][3], 0) + " (kapasitas " + str(int(room_sheet.cell_value(winner[i][3], 1))) + ")"
        worksheet.write(row, 3, room)

        i += 1
        row += 1

    workbook.close()
