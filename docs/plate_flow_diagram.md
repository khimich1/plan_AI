# Поток плит в боте: схема с этапами, оптимизатором и визуализатором

## Полная схема (Mermaid)

Скопируй блок ниже в [mermaid.live](https://mermaid.live) для просмотра и экспорта в PNG/SVG.

```mermaid
flowchart TB
    subgraph этап1["Этап 1: Создание КП"]
        A[Текст/фото списка плит] --> B[commercial]
        B --> C[config_and_data<br/>set_plate_lists_from_text]
        C --> D[order_data / orders_2d]
        D --> O1[core/optimization<br/>optimize_with_cascading_longitudinal_cuts]
        O1 --> O1out[optimization_result<br/>смета, кол-во плит]
        O1out --> B
        B --> V1[core/visualization<br/>visualize_plan]
        V1 --> V1out[Схема раскладки PDF/PNG]
        B --> E[kp_db.save_kp_to_db]
    end

    subgraph БД["БД plita.db"]
        E --> F[(kp_plates<br/>в производстве)]
        G[(plate_rests)]
    end

    subgraph этап2["Этап 2: Планирование"]
        H[production_create<br/>выбор даты, дорожек, КП]
        H --> I[Выбор плит по КП<br/>kp_plate_ids]
        I --> J[production_execution<br/>load_and_plan_production]
    end

    subgraph этап3["Этап 3: Раскладка и треки"]
        F --> J
        G --> J
        J --> J1[find_matching_rests]
        J1 --> J
        J --> O2[core/optimization<br/>optimize_with_cascading_longitudinal_cuts]
        O2 --> O2out[optimization_result<br/>primary_cuts, plate_assignments]
        O2out --> J
        J --> S[core/visualization<br/>split_sequence_into_tracks]
        S --> T[orders_2d + all_tracks_list<br/>state]
    end

    subgraph этап4["Этап 4: Просмотр и экспорт"]
        T --> P[production_day_view<br/>просмотр дня]
        P --> V2[visualize_plan<br/>схема раскладки дня]
        T --> Q[production_export<br/>save_current_plan]
        Q --> Q1[create_gantt_excel<br/>диаграмма Ганта XLSX]
        Q --> M[kp_db.mark_plates_as_planned]
        M --> F
    end

    subgraph этап5["Этап 5: Завершение дня"]
        R[production_completion<br/>завершение дня]
        R --> N[kp_db.move_plates_to_completed]
        N --> F
        N --> P2[(completed_plates)]
    end

    T --> R
```

## Упрощённый линейный поток

```mermaid
flowchart TB
    A[Текст/фото плит] --> B[commercial]
    B --> C[config_and_data]
    C --> D[orders_2d]
    D --> O1[ОПТИМИЗАТОР: смета КП]
    O1 --> B
    B --> V1[ВИЗУАЛИЗАТОР: схема раскладки]
    B --> E[save_kp_to_db]
    E --> F[(kp_plates)]
    F --> G[production_create]
    G --> H[production_execution]
    I[(plate_rests)] -.->|find_matching_rests| H
    H --> O2[ОПТИМИЗАТОР: раскрой]
    O2 --> S[split_sequence_into_tracks]
    S --> T[orders_2d + треки]
    T --> V2[ВИЗУАЛИЗАТОР: visualize_plan / Ганта]
    T --> X[production_export]
    X --> M[mark_plates_as_planned]
    M --> F
    T --> Y[production_completion]
    Y --> N[move_plates_to_completed]
    N --> P[(completed_plates)]
```

## Сводка этапов

| Этап | Handlers | Оптимизатор | Визуализатор | БД |
|------|----------|-------------|--------------|-----|
| **1. КП** | commercial | Да (смета по orders_2d) | visualize_plan — схема раскладки | save_kp_to_db → kp_plates |
| **2. Планирование** | production_create | — | — | чтение kp_plates |
| **3. Раскладка** | production_execution | Да (раскрой → треки) | split_sequence_into_tracks | kp_plates + plate_rests, find_matching_rests |
| **4. Просмотр/экспорт** | production_day_view, production_export | данные из state | visualize_plan, create_gantt_excel | mark_plates_as_planned → kp_plates |
| **5. Завершение дня** | production_completion | — | — | move_plates_to_completed → completed_plates |
