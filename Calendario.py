import csv
import random
from collections import Counter, defaultdict
from itertools import combinations
from pathlib import Path

seed = 77

conference_A = ["A1", "A2", "A3", "A4", "A5", "A6",
                "A7", "A8", "A9", "A10", "A11", "A12"]
conference_B = ["B1", "B2", "B3", "B4", "B5", "B6", 
                "B7", "B8", "B9", "B10", "B11", "B12"]


def check_teams(conf_a, conf_b):
    if len(conf_a) != 12 or len(conf_b) != 12:
        raise ValueError("Devono esserci esattamente 12 squadre in Conference A e 12 in Conference B.")
    teams = list(conf_a) + list(conf_b)
    duplicates = [team for team, count in Counter(teams).items() if count > 1]
    if duplicates:
        raise ValueError(f"Le squadre devono avere nomi unici. Duplicati: {duplicates}")


def andata_intra(teams):
    """
    Round-robin di sola andata per una conference da 12 squadre.
    Produce 11 giornate da 6 partite ciascuna.
    """
    teams = list(teams)
    n = len(teams)
    if n % 2 != 0:
        raise ValueError("Il numero di squadre deve essere pari.")

    rounds = []

    for rnd in range(n - 1):
        coppie = []
        for i in range(n // 2):
            t1 = teams[i]
            t2 = teams[n - i - 1]

            # Alternanza casa/trasferta nella sola andata.
            # Il ritorno verra' poi generato invertendo le partite.
            if i == 0:
                home, away = (t2, t1) if rnd % 2 == 0 else (t1, t2)
            else:
                home, away = (t1, t2) if (rnd + i) % 2 == 0 else (t2, t1)

            coppie.append((home, away))

        rounds.append(coppie)

        # Rotazione corretta: prima squadra fissa, le altre ruotano.
        # Nota: teams[1:-1] deve essere concatenato, non inserito come lista annidata.
        teams = [teams[0]] + [teams[-1]] + teams[1:-1]

    return rounds


def mirror_andata(rounds):
    """Genera il ritorno invertendo casa e trasferta dell'andata."""
    return [[(away, home) for home, away in day] for day in rounds]


def interconf(conferenceA, conferenceB):
    """
    Genera 12 giornate interconference.
    Ogni squadra A gioca una volta contro ogni squadra B.
    In ogni giornata ci sono 12 partite, con 6 squadre A in casa e 6 squadre B in casa.
    Ogni squadra chiude l'interconference con 6 casa e 6 trasferta.
    """
    n = len(conferenceA)
    if n != len(conferenceB):
        raise ValueError("Le conference devono avere la stessa lunghezza.")
    if n != 12:
        raise ValueError("Devono esserci 12 squadre per conference.")

    rounds = []

    for rnd in range(n):
        coppie = []
        for j in range(n):
            ta = conferenceA[j]
            tb = conferenceB[(j + rnd) % n]

            # Pattern bilanciato:
            # - 6 partite con squadra A in casa e 6 con squadra B in casa per ogni giornata;
            # - 6 casa e 6 trasferta per ogni singola squadra sull'interconference.
            if (j + (rnd // 2)) % 2 == 0:
                coppie.append((ta, tb))
            else:
                coppie.append((tb, ta))

        rounds.append(coppie)

    return rounds


def count_home_away(all_rounds):
    stats = defaultdict(lambda: {"home": 0, "away": 0})
    for giornata in all_rounds:
        for home, away in giornata:
            stats[home]["home"] += 1
            stats[away]["away"] += 1
    return stats


def phase_for_day(day_number):
    if 1 <= day_number <= 11:
        return "intraconference_andata"
    if 12 <= day_number <= 23:
        return "interconference"
    if 24 <= day_number <= 34:
        return "intraconference_ritorno"
    raise ValueError(f"Giornata non valida: {day_number}")


def validate_schedule(all_rounds, conf_a, conf_b):
    """
    Controlli:
    1) 34 giornate totali.
    2) Ogni giornata ha 12 partite e ogni squadra gioca una sola volta.
    3) Giornate 1-11 e 24-34 solo intraconference.
    4) Giornate 12-23 solo interconference.
    5) Ogni coppia intra-conference compare due volte, una casa e una fuori.
    6) Ogni coppia interconference compare una volta.
    7) Ogni squadra ha 17 casa e 17 trasferta.
    8) Nell'interconference ogni squadra ha 6 casa e 6 trasferta.
    """
    check_teams(conf_a, conf_b)

    teams = set(conf_a) | set(conf_b)
    set_a = set(conf_a)
    set_b = set(conf_b)

    if len(all_rounds) != 34:
        raise ValueError(f"Numero giornate errato: {len(all_rounds)}. Atteso: 34.")

    if sum(len(g) for g in all_rounds) != 34 * 12:
        raise ValueError("Numero partite totale errato. Attese 408 partite.")

    def same_conference(x, y):
        return (x in set_a and y in set_a) or (x in set_b and y in set_b)

    # 1) validazione giornate + struttura blocchi
    for idx, giornata in enumerate(all_rounds, start=1):
        if len(giornata) != 12:
            raise ValueError(f"Giornata {idx}: numero partite diverso da 12.")

        seen = []
        for home, away in giornata:
            if home == away:
                raise ValueError(f"Giornata {idx}: {home} gioca contro se stessa.")
            if home not in teams or away not in teams:
                raise ValueError(f"Giornata {idx}: squadra non riconosciuta in {home}-{away}.")
            seen.extend([home, away])

        if len(set(seen)) != 24:
            raise ValueError(f"Giornata {idx}: una squadra gioca piu' di una volta.")
        if set(seen) != teams:
            raise ValueError(f"Giornata {idx}: non tutte le squadre sono presenti.")

        if 1 <= idx <= 11 or 24 <= idx <= 34:
            if not all(same_conference(home, away) for home, away in giornata):
                raise ValueError(f"Giornata {idx}: deve contenere solo partite intraconference.")
        elif 12 <= idx <= 23:
            if any(same_conference(home, away) for home, away in giornata):
                raise ValueError(f"Giornata {idx}: deve contenere solo partite interconference.")

    # 2) conteggio accoppiamenti
    intra_counts = defaultdict(list)
    inter_counts = defaultdict(int)
    inter_rounds = all_rounds[11:23]

    for giornata in all_rounds:
        for home, away in giornata:
            if same_conference(home, away):
                intra_counts[frozenset((home, away))].append((home, away))
            else:
                a = home if home in set_a else away
                b = away if away in set_b else home
                inter_counts[(a, b)] += 1

    for conf in (conf_a, conf_b):
        for x, y in combinations(conf, 2):
            matches = intra_counts[frozenset((x, y))]
            if len(matches) != 2:
                raise ValueError(f"Coppia intra {x}-{y}: attese 2 partite, trovate {len(matches)}.")
            homes = {m[0] for m in matches}
            if homes != {x, y}:
                raise ValueError(f"Coppia intra {x}-{y}: casa/trasferta non invertita correttamente.")

    for a in conf_a:
        for b in conf_b:
            if inter_counts[(a, b)] != 1:
                raise ValueError(f"Coppia inter {a}-{b}: attesa 1 partita, trovate {inter_counts[(a, b)]}.")

    # 3) bilancio totale casa/trasferta
    stats = count_home_away(all_rounds)
    for team in teams:
        h = stats[team]["home"]
        a = stats[team]["away"]
        if h != 17 or a != 17:
            raise ValueError(f"{team}: casa={h}, trasferta={a}, atteso 17/17.")

    # 4) bilancio casa/trasferta solo interconference
    inter_stats = count_home_away(inter_rounds)
    for team in teams:
        h = inter_stats[team]["home"]
        a = inter_stats[team]["away"]
        if h != 6 or a != 6:
            raise ValueError(f"{team}: interconference casa={h}, trasferta={a}, atteso 6/6.")

    return True


def gen_schedule(confA, confB, seed=seed, shuffle=True):
    check_teams(confA, confB)

    rng = random.Random(seed)
    confA_work = list(confA)
    confB_work = list(confB)

    if shuffle:
        rng.shuffle(confA_work)
        rng.shuffle(confB_work)

    # Andata intraconference: 11 giornate, 6 partite A + 6 partite B.
    intra_A_andata = andata_intra(confA_work)
    intra_B_andata = andata_intra(confB_work)

    intra_andata = []
    for ra, rb in zip(intra_A_andata, intra_B_andata):
        giornata = ra + rb
        if shuffle:
            rng.shuffle(giornata)
        intra_andata.append(giornata)

    if shuffle:
        rng.shuffle(intra_andata)

    # Interconference: 12 giornate da 12 partite.
    inter = interconf(confA_work, confB_work)
    if shuffle:
        for giornata in inter:
            rng.shuffle(giornata)
        rng.shuffle(inter)

    # Ritorno intraconference: stesso ordine dell'andata, casa/trasferta invertite.
    intra_ritorno = mirror_andata(intra_andata)

    all_rounds = intra_andata + inter + intra_ritorno

    validate_schedule(all_rounds, confA, confB)
    return all_rounds


def schedule_to_rows(all_rounds, conf_a, conf_b):
    set_a = set(conf_a)
    rows = []

    for day_number, giornata in enumerate(all_rounds, start=1):
        fase = phase_for_day(day_number)
        for match_number, (home, away) in enumerate(giornata, start=1):
            rows.append({
                "giornata": day_number,
                "fase": fase,
                "match_giornata": match_number,
                "casa": home,
                "trasferta": away,
                "conference_casa": "A" if home in set_a else "B",
                "conference_trasferta": "A" if away in set_a else "B",
            })

    return rows


def export_csv(all_rounds, conf_a, conf_b, output_path="C:/Users/994944/Desktop/Progetti/calendario_fantabasket.csv"):
    rows = schedule_to_rows(all_rounds, conf_a, conf_b)
    fieldnames = [
        "giornata",
        "fase",
        "match_giornata",
        "casa",
        "trasferta",
        "conference_casa",
        "conference_trasferta",
    ]

    output_path = Path(output_path)
    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    return output_path


def print_schedule(all_rounds):
    for idx, giornata in enumerate(all_rounds, start=1):
        print(f"\n=== Giornata {idx} - {phase_for_day(idx)} ===")
        for home, away in giornata:
            print(f"{home} vs {away}")


def print_home_away_summary(all_rounds):
    stats = count_home_away(all_rounds)
    print("\nRiepilogo casa/trasferta")
    print("Squadra | Casa | Trasferta | Totale")
    print("-" * 36)
    for team in sorted(stats):
        home = stats[team]["home"]
        away = stats[team]["away"]
        print(f"{team:7s} | {home:4d} | {away:9d} | {home + away:6d}")


if __name__ == "__main__":
    calendario = gen_schedule(conference_A, conference_B, seed=seed, shuffle=True)
    csv_path = export_csv(calendario, conference_A, conference_B, "C:/Users/994944/Desktop/Progetti/calendario_fantabasket.csv")
    print("Calendario valido: tutti i vincoli sono rispettati.")
    print(f"Giornate: {len(calendario)}")
    print(f"Partite totali: {sum(len(g) for g in calendario)}")
    print(f"CSV esportato: {csv_path}")
    print_home_away_summary(calendario)
