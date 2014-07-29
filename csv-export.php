<?php
/*
Plugin Name: Excel exporter
Plugin URI:
Description: CSV Export will create a csv to excel file of published  post data.
Version: 1.1
Author: Stratiq "Orginal code by Mindfire Solutions"
Author URI:
*/
ob_end_clean();
ob_start();
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');
require_once('PHPExcel.php');
require_once('ExportHeading.php');
require_once('Printcontent.php');
// to get the value of export button and trigger the export event function for the organizer details.
add_action('init','get_organizer');
function get_organizer() {
	if(isset($_POST['exportevents']) && $_POST['exportevents']) {
		export_organizer_details();
		exit();
	} else {
		return null;
	}
}
global $wpdb;

add_action('admin_enqueue_scripts', 'siq_enqueue_scripts');

function siq_enqueue_scripts() {
	wp_enqueue_script('jquery-export',plugins_url('/js/csv_export.js',__FILE__), array('jquery'));
	wp_enqueue_script('jquery-date',plugins_url('/js/jquery-ui-datepicker.js',__FILE__) , array('jquery', 'jquery-ui-datepicker') );
	wp_enqueue_style('jquery-date-css',plugins_url('/css/datepicker.css',__FILE__));
	wp_enqueue_style('jquery-style','http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css');
}

//function to export organizer details
function export_organizer_details() {
	global $wpdb,$post;
	$start_date='';
	$category = $_POST['siq-organization-category'];
	$start_date = $_POST['startdate'];
	$show_start_date = date('Y-m-d', strtotime($start_date));     		//get the mysql date format  for the starting date
	$end_date = $_POST['enddate'];
	$show_end_date = date('Y-m-d', strtotime($end_date));
	$sitename = sanitize_key( get_bloginfo( 'name' ) )/*  .'-'. sanitize_key( $args['content'] ) .'-' */;

	if(isset($args['start_date']) && isset($args['end_date']) && strlen($args['start_date']) > 4 && strlen($args['end_date']) > 4) {
		$filename = $sitename . $args['start_date'] . '-' . $args['end_date'].'xls';
	} else {
		$filename = $sitename . date( 'Y-m-d' ).'.xls';
	}

	global $wp_query;

	$posts_table = $wpdb->prefix . 'posts';

	//to sort the array according to the headings in the excel sheet
	$post_ids = $wpdb->get_col( "SELECT ID FROM $posts_table where post_type='organization' and DATE_FORMAT($posts_table.post_date,'%Y-%m-%d')  between '$show_start_date' and '$show_end_date' and $posts_table.post_status='publish'" );

	if($post_ids) {

			if (isset($category) && $category != 'all') {
				$posts = $wpdb->get_results(
					"SELECT $posts_table.ID as 'Post ID',DATE_FORMAT($posts_table.post_date,'%b %d %Y') as 'Published Date',$posts_table.post_title as Title
					FROM $posts_table
					LEFT JOIN $wpdb->term_relationships ON
					($posts_table.ID = $wpdb->term_relationships.object_id)
					LEFT JOIN $wpdb->term_taxonomy ON
					($wpdb->term_relationships.term_taxonomy_id = $wpdb->term_taxonomy.term_taxonomy_id)
					WHERE $posts_table.post_type='organization'
					AND $posts_table.post_status='publish'
					AND $wpdb->term_taxonomy.taxonomy = 'org_category'
					AND $wpdb->term_taxonomy.term_id = " . $category . "
					AND DATE_FORMAT($posts_table.post_date,'%Y-%m-%d')  between '$show_start_date' AND '$show_end_date' GROUP BY $posts_table.ID asc", ARRAY_A
				);
			} else {
				$posts = $wpdb->get_results(
					"SELECT p.ID as 'Post ID',DATE_FORMAT(p.post_date,'%b %d %Y') as 'Published Date',p.post_title as Title
					FROM $posts_table p
					WHERE p.post_type='organization'
					AND p.post_status='publish'
					AND DATE_FORMAT(post_date,'%Y-%m-%d')  between '$show_start_date' AND '$show_end_date' GROUP BY p.ID asc", ARRAY_A
				);
			}

			$count=3;
			//for each post id which matches post id in the postmeta table
			foreach($posts as $post) {
					$i=0;
					$contact_multiple =  get_post_meta( $post['Post ID'],'contacts_multi',true );
					if($contact_multiple==0)                          //if first row is empty then i=0
						$i=$contact_multiple;
					else
						$i= $contact_multiple-($contact_multiple-1);    //to get the first row eg 6-(6-1)

					if($i==1)          //if $i==1 we get contacts of 0th row
						 $array_of_meta_keys=array('acronym','description','notes','programs','annual_events','postal_address','street','city','province','postal_code','phone','fax','email','website','open_hours','square_footage','contacts_multi_'.($i-1).'_contact_name','contacts_multi_'.($i-1).'_contact-number','contacts_multi_'.($i-1).'_contact-email','contacts_multi_'.($i-1).'_contact-position','tags','category');
					else
						$array_of_meta_keys=array('acronym','description','notes','programs','annual_events','postal_address','street','city','province',
						'postal_code','phone','fax','email','website','open_hours','square_footage','contacts_multi_contact_name','contacts_multi_contact-number',
						'contacts_multi_contact-email','contacts_multi_contact-position','tags','category');

					$titles = array_merge( array_keys( $post ),$array_of_meta_keys);
					if($count==3) {	  //to print the array from the 3rd row
						$temp=array();
						foreach($titles as $title) {
							$temp[]=$title;
						}
						//renaming the elements of the excel labels
						$replaced_array=array_replace($temp,array(3=>'Acronym',4=>'Description',5=>'Notes',6=>'Programs',7=>'Annual Events',8=>'Postal Address',9=>'Street',10=>'City',11=>'Province/State',12=>'Zip Code',13=>'Phone',14=>'Fax',15=>'Email',16=>'Website',17=>'Open Hours',18=>'Meeting space',19 =>'Contact Name',20=>'Contact Phone',21=>'Contact Email',22=>'Contact Position',23=>'Category or Tag',24=>'Type of Category or Tag'));
						//export the labels in the excel sheet
							ExportToExcel($replaced_array,$filename);
					}
					$empties = array_fill_keys( $titles, '' );
					$post = array_merge( $empties, $post );
				//now comes the content part
				if($i==1)
					$array_of_titles = array( 'acronym','description','notes','programs','annual_events','postal_address',
					'street','city','province','postal_code','phone','fax','email','website','open_hours',
					'square_footage','contacts_multi_'.($i-1).'_contact_name','contacts_multi_'.($i-1).'_contact-number',
					'contacts_multi_'.($i-1).'_contact-email','contacts_multi_'.($i-1).'_contact-position','tags','category');
				else
					$array_of_titles = array( 'acronym','description','notes','programs','annual_events','postal_address',
					'street','city','province','postal_code','phone','fax','email','website','open_hours','square_footage',
					'contacts_multi_contact_name','contacts_multi_contact-number','contacts_multi_contact-email',
					'contacts_multi_contact-position','tags','category');

				$array_title = "'";
				$array_title .= implode("','", $array_of_titles);
				$array_title .= "'";
				$postmeta = $wpdb->get_results( $wpdb->prepare("SELECT * FROM ".$wpdb->prefix."postmeta
																	where post_id = %d and meta_key in
																	(".$array_title.") GROUP BY meta_key", $post['Post ID']), ARRAY_A );
				foreach($postmeta as &$post_meta_val) {
					foreach( $post_meta_val as &$meta_val ) {
						if ( is_string( $meta_val ) )
							$meta_val = strip_tags( $meta_val );
					}
				}
				$post_cats=get_the_terms( $post['Post ID'], 'org_category' );
				$post_tag=get_the_tags($post['Post ID']);
				//get all the post meta to populate in the excel sheet
				if ( $postmeta ) {
					foreach( $postmeta as $meta ) {
							$post[$meta['meta_key']] = $post[$meta['meta_key']].' '. $meta['meta_value'];
					}
				}
				$cont = get_post($post['Post ID'],ARRAY_A);
							 foreach( $cont as &$desc ) {
								if ( is_string( $desc ) )
									$desc = strip_tags( $desc );
							}
							$post['description']= $cont['post_content'];
				//get the tags associated
				$post_tags='';
				if ($post_tag) {
					foreach($post_tag as $tag) {
					$post_tags=$post_tags.$tag->name.',';
					}
					$post_strip_tag=rtrim($post_tags,',');
				  $post['tags']=$post_strip_tag;

				}

				$new_cat = array();
				//get the categories
				if($post_cats) {
					foreach ( $post_cats as $term ) {
						$new_cat[] = $term->name;
					}
				$post_category = join( ", ", $new_cat );
				$post['category']=$post_category;
				}

					print_content($post,$count);
					$count++;


			}//foreach
	}//if
// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="'.$filename.'"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0
		global $objWriter;
		// Write file to the browser
		$objWriter->save('php://output');
    ob_end_flush();
		exit;
}//export_organizer_details

//to add an export button in the edit screen of organization post
add_filter( 'views_edit-organization', 'add_button_to_views' );
function add_button_to_views( $views ) {
	$views['filterbutton1'] = '<form method="POST" action=""><input type="text" class="MyDate" name="startdate" id="startdate" value="Start Date "/>';
	$views['filterbutton2'] = '<input type="text" class="MyDate" name="enddate" id="enddate" value="End Date"/>';



	$categories = get_terms('org_category');
	if(!empty($categories) && is_array($categories)){
		$views['category'] = '<select id="siq-organization-category" name="siq-organization-category">';
		$views['category'] .= '<option value="all">-- All --</option>';
		foreach($categories as $category){
			$views['category'] .= '<option value="' . $category->term_id . '">' . $category->name . '</option>';
		}
		$views['category'] .= '</select>';
	}

	$views['exportbutton'] = '<input type="submit" id="exportevents" name="exportevents" class="button button-primary" value="export"/></form>';

    return $views;
}
?>