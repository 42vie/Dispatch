Interface_ACTION_BLOCKS

Découpage de Interface.html par gros blocs lisibles :
- action_blocks/css/02_MAIN_STYLE_FULL.html = tout le gros style principal
- action_blocks/html = blocs HTML par onglet/modal : Header, Tabs, Dispatch, COMAT, NTO, Options, Modals
- action_blocks/js = blocs JS par action/fonction métier : State, Render Master, Kanban, COMAT, NTO, Options, FC Params, Undo, Dark Mode, etc.
- action_blocks/patches = blocs ajoutés en bas du fichier : Clean Restart, Workflow Transport Chips, EDOC Full, CMR Full UI

index.html reste le fichier complet original.
RECONSTRUCTION_CHECK.json vérifie que les blocs reconstruisent Interface.html sans perte.
