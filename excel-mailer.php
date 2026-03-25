<?php
/*
Plugin Name: CF7 to Excel
Description: Speichert Formulareingaben von Contact Form 7 in einer hochgeladenen Excel-Vorlage.
Version: 1.2
Author: Peter Rader
*/


if (!defined('ABSPATH')) exit;

require_once __DIR__ . '/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// ----------------------------
// Admin-Seite für Mapping + Vorlage
// ----------------------------
add_action('admin_menu', function() {
    add_options_page(
        'CF7 to Excel Mapping',
        'CF7 .ODS Mapping',
        'manage_options',
        'cf7-excel-mapping',
        'cf7_excel_mapping_page'
    );
});

function cf7_form_list_page() {
    if (!current_user_can('manage_options')) return;
    global $wpdb;
    $results = $wpdb->get_results("SELECT * FROM {$wpdb->posts} WHERE post_type LIKE 'wpcf7%'");

    // Prüfen, ob ein Formular ausgewählt wurde
    if (isset($_POST['form_id']) && !empty($_POST['form_id'])) {
        $form_id = intval($_POST['form_id']);
        $form_content = get_post_field('post_content', $form_id);

        // Shortcodes extrahieren (nur Eingabefelder wie text, email, etc.)
        preg_match_all('/\[(text|email|textarea|tel|url|number|checkbox|radio|file)[^\]]*\]/', $form_content, $matches);
        $fields = $matches[0]; // Nur die relevanten Shortcodes extrahieren
    }
    ?>
	<div class="wrap excelmailer">
		<p>Besucher der Internetseite können Formulare ausfüllen und Absenden. Die eingegebenen Daten können Felder in ODS-Dateien befüllen. Die so ausgefüllten ODS-Dateien werden als eMail-Anhang versendet.</p> 
	    <h1>Alle Contact Form 7 Formulare</h1>
        <table>
            <tr>
                <th>Formular:</th>
                <td>
        			<form method="POST">
                        <select name="form_id" onchange="this.form.submit()">
                            <option disabled value="" selected>– Bitte Wählen –</option>
                            <?php if (!empty($results)) : ?>
                                <?php foreach ($results as $form) : ?>
                                    <option value="<?php echo esc_attr($form->ID); ?>" <?php echo (isset($form_id) && $form_id == $form->ID) ? 'selected' : ''; ?>>
                                        <?php echo esc_html($form->post_title); ?>
                                    </option>
                                <?php endforeach; ?>
                            <?php else : ?>
                                <option disabled value="">Es wurden keine Formulare gefunden.</option>
                            <?php endif; ?>
                        </select>
			        </form>
                </td>
            </tr>
        </table>
		<style>
		/* Stil für die Tabelle mit der Klasse "designx" */
		table.designx {
		    width: 100%;
		    border-collapse: collapse;
		    margin: 20px 0;
		    font-family: Arial, sans-serif;
		    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
		}

		table.designx th, table.designx td {
		    padding: 12px 15px;
		    text-align: left;
		    border-bottom: 1px solid #ddd;
		}

		table.designx th {
		    background-color: #4CAF50;
		    color: white;
		    font-size: 16px;
		}

		table.designx td {
		    background-color: #f9f9f9;
		    font-size: 14px;
		    color: #333;
		}

		/* Hover-Effekt für Zeilen */
		table.designx tr:hover {
		    background-color: #f1f1f1;
		}

		/* Alternierende Zeilenfarben */
		table.designx tr:nth-child(even) {
		    background-color: #f2f2f2;
		}

		table.designx tr:nth-child(odd) {
		    background-color: #fff;
		}

		/* Styling für Zellen im letzten Abschnitt */
		table.designx th:last-child, table.designx td:last-child {
		    border-right: 0;
		}

		/* Styling für die erste Spalte */
		table.designx td:first-child {
		    font-weight: bold;
		}

		/* Styling für den Tabellenrahmen */
		table.designx {
		    border: 1px solid #ddd;
		    border-radius: 8px;
		    overflow: hidden;
		}

		/* Responsive Design für kleine Bildschirme */
		@media screen and (max-width: 768px) {
		    table.designx {
		        width: 100%;
		        border: 0;
		    }
		    table.designx th, table.designx td {
		        display: block;
		        width: 100%;
		        box-sizing: border-box;
		    }
		    table.designx td {
		        padding: 10px;
		        text-align: right;
		        border-bottom: 1px solid #ddd;
		    }
		    table.designx th {
		        display: none;
		    }
		}
		table.designx input {width:100%;}
		
		
		.excelmailer p {
			display: block;
			max-width: 400px;
		}

		/* Stil für Buttons */
		.excelmailer button {
		    background-color: #4CAF50; /* Grüner Hintergrund */
		    color: white; /* Weiße Schrift */
		    padding: 10px 20px; /* Innenabstand */
		    font-size: 16px; /* Schriftgröße */
		    border: none; /* Kein Rand */
		    border-radius: 5px; /* Abgerundete Ecken */
		    cursor: pointer; /* Zeiger-Cursor */
		    transition: background-color 0.3s ease; /* Übergang für Hintergrundfarbe */
		}

		.excelmailer button:hover {
		    background-color: #45a049; /* Dunkleres Grün beim Hover */
		}

		.excelmailer button:focus {
		    outline: none; /* Entfernen des Fokusrands */
		}
		
		</style>
        <?php if (isset($fields)) : ?>
			<table width="100%" class="designx">
				<thead>
					<tr>
						<th>Eingabe Feld</th>
						<th>Tabellenreiter</th>
						<th>Tabellenspalte</th>
						<th>Tabellenzeile</th>
					</tr>
				</thead>
				<tbody>
	                <?php foreach ($fields as $field) : ?>
						<tr>
		                    <td><?php echo esc_html($field); ?></td>
							<td><input type="number" min="1" max="20" placeholder="(1 für den ersten Tab)"/></td>
							<td><input type="text" placeholder="(A für die erste Spalte ...)"/></td>
							<td><input type="number" min="1" max="60000" placeholder="(1 für die erste Zeile)"/></td>
						</tr>
	                <?php endforeach; ?>
				</tbody>
            </table>
			<button>Feld-Mapping Speichern</button>
			<button>ODS-Vorlage Hochladen</button>
			<button>ODS-Vorlage Herunterladen</button>
        <?php endif; ?>
    </div>
    <?php
}


function cf7_excel_mapping_page() {
    if (!current_user_can('manage_options')) return;

	
    // Speichern
    if (isset($_POST['cf7_excel_mapping'])) {
        check_admin_referer('cf7_excel_mapping_save', 'cf7_excel_mapping_nonce');
        update_option('cf7_excel_field_mapping', $_POST['cf7_excel_mapping']);

        // Datei-Upload
        if (!empty($_FILES['cf7_excel_template']['name'])) {
            $uploaded = wp_handle_upload($_FILES['cf7_excel_template'], ['test_form' => false]);
            if (isset($uploaded['file'])) {
                update_option('cf7_excel_template_file', $uploaded['file']);
                echo '<div class="updated"><p>Template erfolgreich hochgeladen!</p></div>';
            } else {
                echo '<div class="error"><p>Fehler beim Hochladen der Datei.</p></div>';
            }
        } else {
            echo '<div class="updated"><p>Mapping gespeichert!</p></div>';
        }
    }

    $mapping = get_option('cf7_excel_field_mapping', []);
    $template_file = get_option('cf7_excel_template_file', '');
	
	
	// Holen der CF7 Formulare
    cf7_form_list_page()
    ?>
	
	
	
    <?php
}

// ----------------------------
// CF7 Aktion zum Speichern
// ----------------------------
add_action('wpcf7_mail_sent', 'cf7_to_excel_save_data');

function cf7_to_excel_save_data($contact_form) {
    $submission = WPCF7_Submission::get_instance();
    if (!$submission) return;

    $form_data = $submission->get_posted_data();

    // Vorlage laden
    $template_file = get_option('cf7_excel_template_file');
    if (!$template_file || !file_exists($template_file)) return;

    // Mapping laden
    $field_mapping = get_option('cf7_excel_field_mapping', []);
    if (empty($field_mapping)) return;

    // Excel öffnen
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($template_file);
    $sheet = $spreadsheet->getActiveSheet();
    $row = $sheet->getHighestRow() + 1;

    // Daten eintragen
    foreach ($field_mapping as $field => $column) {
        $sheet->setCellValue($column . $row, isset($form_data[$field]) ? $form_data[$field] : '');
    }

    // Gespeicherte Datei – optional könnte hier ein eigenes Verzeichnis pro Formular angelegt werden
    $save_path = plugin_dir_path(__FILE__) . 'form_data_filled.xlsx';
    $writer = new Xlsx($spreadsheet);
    $writer->save($save_path);
}