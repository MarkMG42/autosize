# autosize — инструкции Codex

Репозиторий подготовлен для разработки автоматизации размеров в SolidWorks 2025.

- [AGENTS.md](AGENTS.md) — точка входа нового чата.
- [engineering-knowledge/AGENTS.md](engineering-knowledge/AGENTS.md) — правила координатора.
- [.codex/config.toml](.codex/config.toml) — три роли: Coder, API Expert, Debugger.
- [engineering-knowledge/agents](engineering-knowledge/agents/) — инструкции подагентов.
- [engineering-knowledge/topics/agent-setup.md](engineering-knowledge/topics/agent-setup.md) — когда привлекать агентов.
- [engineering-knowledge/CHAT_START.md](engineering-knowledge/CHAT_START.md) — промпт для нового чата.

Не считать, что код, архивы или инженерная база репозитория Macro автоматически
доступны здесь.

## Макросы размеров рамы

- [FrameTubeDimensionTrial](FrameTubeDimensionTrial.README.md) — проверка одной
  продольной кромки профильной трубы возле скругления.
- [FrameCutListDimensions](FrameCutListDimensions.README.md) — массовая расстановка
  по телам выбранной детали рамы; швеллеры и уголки проверены, для пропущенных труб
  подготовлена отдельная проба выше.
- [FrameCutListDimensionTrial](FrameCutListDimensionTrial.README.md) — подтверждённая
  одиночная проба размера и размещения в изометрическом виде.
- [FrameCutListProbe](FrameCutListProbe.README.md) — диагностика списка вырезов
  и выбранного компонента без создания размеров.

Старый `AutoDimensionFrame` 2.4.3 сохранён как предыдущая реализация; для уточнённой
задачи «только тела выбранной детали рамы» используйте новые модули выше.
