#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для просмотра данных о стоимости сырья и производственных расходов
"""
import sys
from core.raw_material_db import get_raw_material_cost, get_all_costs, get_statistics

def main():
    if len(sys.argv) > 1:
        # Поиск конкретной плиты
        plate_name = ' '.join(sys.argv[1:])
        cost = get_raw_material_cost(plate_name)
        
        if cost:
            print(f"\n{plate_name}")
            print(f"Stoimost syrya i proizvodstvennyh rashodov: {cost:.2f} rub.")
        else:
            print(f"\nPlita '{plate_name}' ne naydena v baze dannyh")
            print("\nDostupnye plity:")
            all_costs = get_all_costs()
            # Показываем похожие
            similar = [p for p in all_costs.keys() if plate_name.lower() in p.lower()]
            if similar:
                for p in similar[:5]:
                    print(f"  {p}: {all_costs[p]:.2f} rub.")
            else:
                print("  Pohozhie ne naydeny")
    else:
        # Показываем общую статистику и примеры
        print("\n" + "="*70)
        print("  STOIMOST SYRYA I PROIZVODSTVENNYH RASHODOV")
        print("="*70)
        
        stats = get_statistics()
        print(f"\nSTATISTIKA:")
        print(f"  Vsego zapisey: {stats['count']}")
        print(f"  Minimalnaya stoimost: {stats['min']:.2f} rub.")
        print(f"  Maksimalnaya stoimost: {stats['max']:.2f} rub.")
        print(f"  Srednyaya stoimost: {stats['avg']:.2f} rub.")
        
        all_costs = get_all_costs()
        sorted_costs = sorted(all_costs.items(), key=lambda x: x[0])
        
        print(f"\nPERVYE 20 ZAPISEY:")
        for plate, cost in sorted_costs[:20]:
            print(f"  {plate:20s} {cost:10.2f} rub.")
        
        if len(sorted_costs) > 20:
            print(f"\n  ... i eshe {len(sorted_costs) - 20} zapisey")
        
        print(f"\nDlya poiska konkretnoy plity ispolzuyte:")
        print(f"  python view_raw_material_costs.py PB 17-12-6")
        print("="*70 + "\n")

if __name__ == '__main__':
    main()

