# -*- coding: utf-8 -*-
"""
Dieses Modul definiert die RuleEngine.

Die RuleEngine ist verantwortlich für die dynamische Erzeugung von Datenfeldern
basierend auf den in der Projekt-Konfiguration definierten 'generation_rules'.
Sie entkoppelt die Geschäftslogik von der festen Implementierung und unterstützt
verschachtelte Regeln durch topologisches Sortieren der Abhängigkeiten.
"""

import pandas as pd
from collections import deque


class RuleEngine:
    """
    Führt definierte Regeln aus, um Datenzeilen zu verarbeiten und neue
    Felder zu generieren.
    """
    def __init__(self, rules: dict):
        """
        Initialisiert die RuleEngine mit einem Set von Regeln.
        """
        self.rules = rules
        self.sorted_rules = self._get_execution_order()

    def process_row(self, row_data: dict) -> dict:
        """
        Wendet alle definierten Regeln in der korrekten, abhängigkeitsbasierten
        Reihenfolge auf eine einzelne Datenzeile an.
        """
        if not self.sorted_rules: # Falls ein Fehler aufgetreten ist (z.B. Zirkelbezug)
            return {}

        # Wir arbeiten auf einer Kopie, um die generierten Werte für
        # nachfolgende Regeln verfügbar zu machen.
        local_row_data = row_data.copy()
        generated_data = {}

        for target_field in self.sorted_rules:
            rule = self.rules[target_field]
            rule_type = rule.get("type")
            result = ""

            if rule_type == "prioritized_list":
                result = self._apply_prioritized_list(rule, local_row_data)
            elif rule_type == "combine":
                result = self._apply_combine(rule, local_row_data)
            elif rule_type == "conditional":
                result = self._apply_conditional(rule, local_row_data)
            
            generated_data[target_field] = result
            # WICHTIG: Mache das Ergebnis für die nächste Regel verfügbar!
            local_row_data[target_field] = result
            
        return generated_data

    def _get_execution_order(self) -> list:
        """
        Erstellt einen Abhängigkeitsgraphen und sortiert ihn topologisch,
        um die korrekte Ausführungsreihenfolge der Regeln zu ermitteln.
        """
        # 1. Graphen und "in-degree" (Anzahl eingehender Kanten) initialisieren
        graph = {name: [] for name in self.rules}
        in_degree = {name: 0 for name in self.rules}
        
        # 2. Graphen aufbauen
        for target_field, rule in self.rules.items():
            sources = []
            if rule.get("type") in ["prioritized_list", "combine"]:
                sources = rule.get("sources", [])
            elif rule.get("type") == "conditional":
                sources.append(rule.get("if", {}).get("source"))
                sources.append(rule.get("then", {}).get("source"))
                sources.append(rule.get("else", {}).get("source"))

            for source in filter(None, sources):
                # Wenn eine Quelle selbst eine Regel ist, füge eine Kante hinzu
                if source in self.rules:
                    graph[source].append(target_field)
                    in_degree[target_field] += 1
        
        # 3. Topologisches Sortieren (Kahn's Algorithm)
        queue = deque([name for name, degree in in_degree.items() if degree == 0])
        sorted_order = []
        
        while queue:
            node = queue.popleft()
            sorted_order.append(node)
            
            for neighbor in graph[node]:
                in_degree[neighbor] -= 1
                if in_degree[neighbor] == 0:
                    queue.append(neighbor)
        
        # 4. Zirkelbezug prüfen
        if len(sorted_order) == len(self.rules):
            return sorted_order
        else:
            print("FEHLER: Zirkelbezug in den Setzregeln erkannt! Bitte prüfen Sie Ihre Konfiguration.")
            # Identifiziere die Knoten, die Teil des Zyklus sind
            cycle_nodes = {name for name, degree in in_degree.items() if degree > 0}
            print(f"  -> Beteiligte Felder: {', '.join(cycle_nodes)}")
            return None

    # Die apply-Methoden bleiben gleich, sie arbeiten jetzt nur mit den
    # potenziell angereicherten Daten in local_row_data.
    def _apply_prioritized_list(self, rule: dict, row_data: dict) -> str:
        for source_field in rule.get("sources", []):
            value = row_data.get(source_field)
            if value and pd.notna(value): return str(value).strip()
        return ""

    def _apply_combine(self, rule: dict, row_data: dict) -> str:
        parts = []
        for source_field in rule.get("sources", []):
            value = row_data.get(source_field)
            if value and pd.notna(value): parts.append(str(value).strip())
        separator = rule.get("separator", " ").replace("\\n", "\n")
        return separator.join(parts)

    def _apply_conditional(self, rule: dict, row_data: dict) -> str:
        if_clause = rule.get("if", {}); then_clause = rule.get("then", {}); else_clause = rule.get("else", {})
        source_value = str(row_data.get(if_clause.get("source"), "")).strip()
        condition_met = self._evaluate_condition(source_value, if_clause.get("operator"), if_clause.get("value", ""))
        result_source = then_clause.get("source") if condition_met else else_clause.get("source")
        final_value = row_data.get(result_source)
        return str(final_value).strip() if final_value and pd.notna(final_value) else ""

    def _evaluate_condition(self, source_value: str, operator: str, compare_value: str) -> bool:
        if operator == "is_empty": return not source_value
        if operator == "is_not_empty": return bool(source_value)
        compare_list = [v.strip() for v in compare_value.split(';')]
        if operator == "is": return source_value in compare_list
        if operator == "is_not": return source_value not in compare_list
        if operator == "contains": return any(sub in source_value for sub in compare_list)
        return False
