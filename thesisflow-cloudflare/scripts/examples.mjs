const example = {
 student: 'Moritz Hoffmann', matriculation: '1234567', degree: 'B.Sc. Wirtschaftsinformatik',
 department: 'Fachbereich MND', title: 'Konzeption eines digitalen Abschlussarbeiten-Portals',
 question: 'Wie kann ein Webportal den Anmelde- und Freigabeprozess von Abschlussarbeiten im Fachbereich MND vereinfachen?',
 context: 'Die Anmeldung erfolgt bislang über Formulare und E-Mails. Fehlende Angaben und manuelle Freigaben führen zu Rückfragen und schwer nachvollziehbaren Bearbeitungsständen.',
 goal: 'Ein Soll-Prozess mit Rollen, digitaler Themenfreigabe und transparenten Statusanzeigen wird konzipiert und durch einen anklickbaren Prototyp veranschaulicht.',
 method: 'Ist-Prozess und Formulare analysieren, Anforderungen mit User Stories beschreiben, einen Prototyp entwickeln und anhand ausgewählter Nutzungsszenarien prüfen.',
 scope: 'Betrachtet werden Anmeldung und Themenfreigabe. Benotung, Plagiatsprüfung und produktive Hochschulanbindung sind nicht Bestandteil der Arbeit.',
 deliverables: 'Prozessmodell, spezifizierte User Stories, UI-Prototyp und dokumentierte Evaluation.',
 secondExaminer: 'Prof. Dr. Lara Neumann (fiktiv)', plannedStart: '2026-10-15', complete: true, authorized: true
};
export const topics = [
 [1, 'weber', example],
 [2, 'weber', {...example, student: 'Tarek Yilmaz', matriculation: '7654321', degree: 'M.Sc. Wirtschaftsinformatik', title: 'Digitale Freigabeprozesse im Unternehmen', authorized: false}],
 [3, 'klein', {...example, student: 'Lea Sommer', title: 'Visualisierung betrieblicher Kennzahlen'}]
];
