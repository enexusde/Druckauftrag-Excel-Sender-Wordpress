<?php
/*
Plugin Name: CF7 to Excel
Description: Speichert Formulareingaben von Contact Form 7 in einer hochgeladenen Excel-Vorlage.
Version: XXX
Author: Peter Rader
*/


if (!defined('ABSPATH')) exit;

require_once __DIR__ . '/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Ods;

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
    
    if (isset($_POST['form_id']) && !empty($_POST['form_id'])) {
        $form_id = intval($_POST['form_id']);
        $form_content = get_post_field('post_content', $form_id);

        preg_match_all('/\[(text|email|textarea|tel|url|number|checkbox|radio|file)[^\]]*\]/', $form_content, $matches);
        $fields = $matches[0];

        $mapping_data = get_option('cf7_excel_field_mapping'.$form_id, []);
		$filename = get_option('cf7_uploaded_file_name'.$form_id, []);
    }
    ?>
    <div class="wrap excelmailer">
        <p>Besucher der Internetseite können Formulare ausfüllen und Absenden. Die eingegebenen Daten können Felder in .ods-Dateien befüllen. Die so ausgefüllten .ods-Dateien werden als eMail-Anhang versendet.</p> 
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

        <?php if (isset($fields)) : ?>
            <form method="POST">
                <?php wp_nonce_field('cf7_excel_mapping_save', 'cf7_excel_mapping_nonce'); ?>
				<input type="hidden" name="form_id" value="<?php echo $form_id ?>" />
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
                        <?php foreach ($fields as $index => $field) : ?>
                            <tr>
                                <td><?php echo esc_html($field); ?></td>
                                <td>
                                    <input type="number" name="cf7_mapping[<?php echo $index; ?>][tab]" 
                                           min="1" max="20" 
                                           placeholder="(1 für den ersten Tab)"
                                           value="<?php echo isset($mapping_data[$index]['tab']) ? esc_attr($mapping_data[$index]['tab']) : ''; ?>" />
                                </td>
                                <td>
                                    <input type="text" name="cf7_mapping[<?php echo $index; ?>][column]" 
                                           placeholder="(A für die erste Spalte ...)"
                                           value="<?php echo isset($mapping_data[$index]['column']) ? esc_attr($mapping_data[$index]['column']) : ''; ?>" />
                                </td>
                                <td>
                                    <input type="number" name="cf7_mapping[<?php echo $index; ?>][row]" 
                                           min="1" max="60000" 
                                           placeholder="(1 für die erste Zeile)"
                                           value="<?php echo isset($mapping_data[$index]['row']) ? esc_attr($mapping_data[$index]['row']) : ''; ?>" />
                                </td>
                            </tr>
                        <?php endforeach; ?>
						<?php if (empty($fields)) : ?>
						    <tr>
						        <td colspan="4" style="text-align:center; color:#999;">
						            Keine Formular-Felder erkannt.
						        </td>
						    </tr>
						<?php endif; ?>
                    </tbody>
                </table>
                
                <button type="submit" name="cf7_excel_mapping" class="button button-primary">Feld-Mapping Speichern</button>
            </form>
			<form method="POST" enctype="multipart/form-data">
				<input type="hidden" name="form_id" value="<?php echo $form_id ?>" />
				<?php wp_nonce_field('cf7_excel_mapping_save', 'cf7_excel_mapping_nonce'); ?>
				<input type="file" style="display:none" id="upltfods" accept=".ods,.xls,.xlsx" onchange="document.getElementById('upltfodsfn').value=files[0].name;document.getElementById('upltfodsok').click();" name="template"/>
				<input type="hidden" name="uploaded_file_name" id="upltfodsfn">
				<button type="submit" id="upltfodsok" style="display:none" />
				<input type="hidden" name="post_id" value="<?php echo $form_id ?>" />
			</form>
			<button onclick="document.getElementById('upltfods').click();" class="button button-primary">.ods-Datei ändern ...</button>
			<p>Aktuelles Template: <tt><?php echo isset($filename) ? esc_attr($filename) : 'Keine Vorlage hochgeladen!'; ?></tt>
        <?php endif; ?>
    </div>
<?php
}



add_action('admin_enqueue_scripts', function($hook) {
    if ($hook !== 'settings_page_cf7-excel-mapping') return; 
    wp_enqueue_style('cf7-excel-css', plugin_dir_url(__FILE__) . 'excel-mailer.css', [], '1.0');
});

    ?>

    <?php
function cf7_excel_mapping_page() {
    if (!current_user_can('manage_options')) return;
    if (isset($_POST['cf7_excel_mapping']) && check_admin_referer('cf7_excel_mapping_save', 'cf7_excel_mapping_nonce')) {
        $mapping_data = $_POST['cf7_mapping'];
        if (!empty($mapping_data)) {
			$form_id = intval($_POST['form_id']);
            update_option('cf7_excel_field_mapping'.$form_id, $mapping_data);
            echo '<div class="updated"><p>Mapping erfolgreich gespeichert!</p></div>';
        } else {
            echo '<div class="error"><p>Es wurden keine Mapping-Daten übermittelt!</p></div>';
        }
    }
	if (isset($_FILES['template']) && !empty($_FILES['template']['name'])) {
	    $file = $_FILES['template'];
	    $allowed_extensions = ['ods', 'xls', 'xlsx'];
	    $file_extension = pathinfo($file['name'], PATHINFO_EXTENSION);
	    if (in_array($file_extension, $allowed_extensions)) {
	        $new_filename = 'cf7_template_' . time() . '.' . $file_extension;
	        $plugin_dir = plugin_dir_path(__FILE__); // Gibt den absoluten Pfad des Plugin-Verzeichnisses zurück
	        $upload_path = $plugin_dir . $new_filename; // Neuen Pfad zum Speichern der Datei
	        if (move_uploaded_file($file['tmp_name'], $upload_path)) {
				$form_id = intval($_POST['form_id']);
	            update_option('cf7_excel_template_file'.$form_id, $upload_path);
	            $uploaded_file_name = isset($_POST['uploaded_file_name']) ? sanitize_text_field($_POST['uploaded_file_name']) : '';
	            update_option('cf7_uploaded_file_name'.$form_id, $uploaded_file_name);
	            echo '<div class="updated"><p>Datei erfolgreich hochgeladen und gespeichert: ' . esc_html($uploaded_file_name) . '</p></div>';
	        } else {
	            echo '<div class="error"><p>Fehler beim Hochladen der Datei!</p></div>';
	        }
	    } else {
	        echo '<div class="error"><p>Ungültige Dateierweiterung. Bitte lade eine .ods, .xls oder .xlsx Datei hoch.</p></div>';
	    }
	}
    $mapping = get_option('cf7_excel_field_mapping', []);
    $template_file = get_option('cf7_excel_template_file', '');
    cf7_form_list_page();
}



add_filter('wpcf7_mail_components', 'cf7_attach_ods_to_email', 10, 3);

function cf7_attach_ods_to_email($components, $contact_form, $abort) {
    $submission = WPCF7_Submission::get_instance();
    if (!$submission) return $components;

    $form_id = $contact_form->id();
    $form_data = $submission->get_posted_data();

	$template_file = get_option('cf7_excel_template_file' . $form_id); 
	$attachment_name = get_option('cf7_uploaded_file_name' . $form_id); 
    $field_mapping = get_option('cf7_excel_field_mapping' . $form_id, []);

    if (!$template_file || !file_exists($template_file) || empty($field_mapping)) {
        return $components;
    }

    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($template_file);

    foreach ($field_mapping as $index => $mapping) {
        $column = strtoupper(trim($mapping['column'] ?? ''));
        $row = intval($mapping['row'] ?? 0);
        $sheet_index = max(0, intval($mapping['tab'] ?? 1) - 1);
        $sheet = $spreadsheet->getSheet($sheet_index);

        $field_name = array_keys($form_data)[$index] ?? null;
        if ($column && $row > 0 && $field_name && isset($form_data[$field_name])) {
            $sheet->setCellValue($column . $row, $form_data[$field_name]);
        }
    }

    // Datei speichern
    $upload_dir = wp_upload_dir();
    $filename = 'cf7_form_' . $form_id . '_' . time() . '.ods';
    $save_path = trailingslashit($upload_dir['basedir']) . $filename;

    $writer = new Ods($spreadsheet);
    $writer->save($save_path);

    // Datei als Anhang hinzufügen
    if (!isset($components['attachments'])) {
        $components['attachments'] = [];
    }
    $components['attachments'][] = $save_path;

    return $components;
}