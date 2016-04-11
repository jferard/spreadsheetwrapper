/*******************************************************************************
 *     SpreadsheetWrapper - An abstraction layer over some APIs for Excel or Calc
 *     Copyright (C) 2015  J. Férard
 *
 *     This program is free software: you can redistribute it and/or modify
 *     it under the terms of the GNU General Public License as published by
 *     the Free Software Foundation, either version 3 of the License, or
 *     (at your option) any later version.
 *
 *     This program is distributed in the hope that it will be useful,
 *     but WITHOUT ANY WARRANTY; without even the implied warranty of
 *     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *     GNU General Public License for more details.
 *
 *     You should have received a copy of the GNU General Public License
 *     along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *******************************************************************************/
package com.github.jferard.spreadsheetwrapper;

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * The class AbstractSpreadsheetReaderTest tests reading an example file. Works
 * for every wrapper.
 *
 */
public abstract class AbstractSpreadsheetReaderTest { // NOPMD by Julien on
	// 27/11/15 20:34
	/** content of cell A1 */
	private static final String CELL_A1_CONTENT = "OBJECTID,N,10,0";

	/** content of cell A2 */
	private static final int CELL_A2_CONTENT = 170;

	/** content of col C (= 3) */
	private static final List<String> COL_C_CONTENTS = Arrays.asList(
			"ETIQUETTE,C,254", "Immeuble dit « Hôtel Lefèvre »", "Gare",
			"Hôtel de Solas", "Portail", "Tour des Pins", "Hôtel de Fourques",
			"Hôtel de Bénézet", "Hôtel de Guidais", "Hôtel Haguenot", "Maison",
			"Ancien Logis du Chapeau rouge", "Hôtel Bardy", "Immeuble",
			"Ancien bureau d’octroi du Pont Juvénal", "Hôtel Verchant",
			"Hôtel de Querelles", "Hôtel de Cambacérès-Murles",
			"Hôtel de Ricard", "Hôtel de Campan", "Hôtel Montcalm", "Immeuble",
			"Hôtel de Lunas", "Hôtel des Carcassonne ou Hôtel de Gayon",
			"Jardin des plantes",
			"Ancien couvent des Ursulines, anciennes prisons (ex-caserne Grossetti)",
			"Ancienne prison", "Eglise des Pénitents Blancs",
			"Pont-aqueduc dit « Arceaux sur la Lironde »", "Hôtel Lamouroux",
			"Ancien cinéma Pathé", "Halle Castellane",
			"Vestiges de l'ancien ensemble cultuel hébraïque",
			"Temple protestant", "Faculté de Médecine et Musée d’anatomie",
			"Mas de Bagnères", "Ancienne église de Montels",
			"Ancien prieuré de Saint Pierre de  Montaubérou",
			"Domaine du château de Bonnier de la Mosson", "Citadelle",
			"Hôtel Bachy-du-Cayla", "Hôtel de Fizes",
			"Eglise Saint François de la Pierre Rouge, de l'Enclos Saint François",
			"Ancienne église des Cordeliers", "Evêché",
			"Hôtel Hortolès ou Ginestous", "Aqueduc des Arceaux ou de Pitot",
			"Fontaine des Licornes", "Porte de la Blanquerie",
			"Hôtel de Baudon et Mauny", "Fontaine de la Préfecture",
			"Hôtel de Fesquet", "Hôtel de Manse", "Hôtel d’Aurès",
			"Hôpital Général et Clinique Saint Charles",
			"Faculté de Médecine et Musée d’anatomie", "Ancien Hôtel de Ganges",
			"Fontaine des Trois Grâces",
			"Ensemble monumental de la Promenade du Peyrou",
			"Cathédrale Saint Pierre", "Eglise Sainte Croix de Celleneuve",
			"Eglise Notre-Dame-des-Tables", "Eglise Saint-Denis",
			"Eglise Sainte-Eulalie", "Hôtel de Mirman",
			"Chapelle de l'ancien couvent des Récollets",
			"hôtel de Grave, hôtel de Villarmois, hôtel de Noailles",
			"Hôtel Rey", "Hôtel de Sarret dit « de la Coquille »",
			"Hôtel des Trésoriers de la Bourse", "Hôtel de Varennes",
			"Vestiges de l'ancien ensemble cultuel hébraïque",
			"Jardins de la reine, Ancien Rectorat et Intendance Jardin des Plantes",
			"Ancien Hôtel de Belleval", "Hôtel de Joubert", "Hôtel d’Uston",
			"Ancienne maison ou couvent de la Miséricorde",
			"Hôtel de Bernard Duffau et ancien grand séminaire", "Hôtel Périer",
			"Hôtel de Boussugues", "Hôtel Pas de Beaulieu", "Hôtel de Castries",
			"Immeuble", "Ancienne église de la Visitation",
			"Eglise de Castelnau-le-Lez",
			"Eglise paroissiale Sainte Thérèse de Lisieux", "Hôtel de Claris",
			"Hôtel des Trésoriers de France (ou Hôtel du Lunaret)",
			"Escalier monumental", "Hôtel", "Hôtel Hostalier",
			"Hôtel de Saint Côme", "Hôtel de Magny ou Cabanes de Puimisson",
			"Hôtel de Castan", "Palais de Justice", "Château de Flaugergues",
			"Château Levat", "Château de la Mogère",
			"Château et Parc de la Piscine",
			"Ancien observatoire dit «Tour de la Babotte »",
			"Hôtel de Montferrier", "Hôtel de Griffy", "Hôtel Estorc",
			"Hôtel de la Société Royale des Sciences", "Hôtel des Vignes",
			"Hôtel Pomier-Layrargues", "Hôtel d'Hostalier",
			"Palais des Rois d’Aragon", "Hôtel de Roquemaure", "Hôtel d’Avèze",
			"Hôtel de Beaulac", "Hôtel Deydé", "Hôpital Général",
			"Hôtel de Saint-Félix", "Hôtel Lecourt", "Château d’Ô",
			"Collège des Ecossais");

	/** content of row 2 */
	private static final List<Object> ROW_2_CONTENTS = Arrays.<Object> asList(
			Integer.valueOf(170), "Monuments historiques classés ou inscrits",
			"Immeuble dit « Hôtel Lefèvre »", "inscrit Inv. MH 19 11 1985",
			"En totalité", Double.valueOf(333.810638315),
			Double.valueOf(93.4774612976));

	/** name of the sheet */
	private static final String SHEET_NAME = "Feuille1";

	/** name of the test */
	@Rule
	public TestName name = new TestName();

	/** simple logger : static initialization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** document reader */
	protected SpreadsheetDocumentReader documentReader;

	/** factory */
	protected SpreadsheetDocumentFactory factory;

	/** reader */
	protected SpreadsheetReader sheetReader;

	/** set the test up */
	@Before
	@SuppressWarnings("nullness")
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL resourceURL = this.getProperties().getResourceURL();
			Assume.assumeNotNull(resourceURL);

			final InputStream inputStream = resourceURL.openStream();
			this.documentReader = this.factory.openForRead(inputStream);
			Assert.assertEquals(1, this.documentReader.getSheetCount());
			this.sheetReader = this.documentReader.getSpreadsheet(0);
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** clean after test */
	@After
	public void tearDown() {
		try {
			this.documentReader.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** read the content of cell A1 */
	@Test
	public final void testCellA1Content() {
		try {
			Assert.assertEquals(AbstractSpreadsheetReaderTest.CELL_A1_CONTENT,
					this.sheetReader.getCellContent(0, 0));
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}

	/** read the text of cell A1 */
	@Test
	public final void testCellA1IsAText() {
		try {
			Assert.assertEquals(AbstractSpreadsheetReaderTest.CELL_A1_CONTENT,
					this.sheetReader.getText(0, 0));
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}

	/**
	 * read the content of cell A1 as boolean : must return null
	 */
	@Test
	public final void testCellA1IsNotABoolean() {
		Assert.assertNull(this.sheetReader.getBoolean(0, 0));
	}

	/**
	 * read the content of cell A1 as date : must return null
	 */
	@Test
	public final void testCellA1IsNotADate() {
		Assert.assertNull(this.sheetReader.getDate(0, 0));
	}

	/**
	 * read the content of cell A1 as double : must return null
	 */
	@Test
	public final void testCellA1IsNotADouble() {
		Assert.assertNull(this.sheetReader.getDouble(0, 0));
	}

	/**
	 * read the content of cell A1 as formula : must return null
	 */
	@Test
	public final void testCellA1IsNotAFormula() {
		Assert.assertNull(this.sheetReader.getFormula(0, 0));
	}

	/**
	 * read the content of cell A1 as integer : must return null
	 */
	@Test
	public final void testCellA1IsNotAnInteger() {
		Assert.assertNull(this.sheetReader.getInteger(0, 0));
	}

	/** read the content of cell A2 as double */
	@Test
	public final void testCellA2AsADouble() {
		try {
			Assert.assertEquals(
					Double.valueOf(
							AbstractSpreadsheetReaderTest.CELL_A2_CONTENT),
					this.sheetReader.getDouble(1, 0));
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}

	/** read the content of cell A2 as integer */
	@Test
	public final void testCellA2AsAnInteger() {
		try {
			Assert.assertEquals(
					Integer.valueOf(
							AbstractSpreadsheetReaderTest.CELL_A2_CONTENT),
					this.sheetReader.getInteger(1, 0));
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}

	/** read the content of cell A2. */
	@Test
	public final void testCellA2Content() {
		try {
			Assert.assertEquals(AbstractSpreadsheetReaderTest.CELL_A2_CONTENT,
					this.sheetReader.getCellContent(1, 0));
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}

	/**
	 * read the content of cell A1 as formula : must return null
	 */
	@Test
	public final void testCellA2IsNotAFormula() {
		Assert.assertNull(this.sheetReader.getFormula(1, 0));
	}

	/**
	 * read the content of cell A1 as text : must return null
	 */
	@Test
	public final void testCellA2IsNotAText() {
		Assert.assertNull(this.sheetReader.getText(1, 0));
	}

	/**
	 * read the content of cell A1 as boolean : must return null
	 */
	@Test
	public final void testCellDateA2IsNotABoolean() {
		Assert.assertNull(this.sheetReader.getBoolean(1, 0));
	}

	/**
	 * read the content of cell A1 as date : must return null
	 */
	@Test
	public final void testCellDateA2IsNotADate() {
		Assert.assertNull(this.sheetReader.getDate(1, 0));
	}

	/** read the content of col contents */
	@Test
	public final void testGetColContents() {
		try {
			Assert.assertEquals(AbstractSpreadsheetReaderTest.COL_C_CONTENTS,
					this.sheetReader.getColContents(2));
		} catch (final IllegalArgumentException e) {
			Assert.fail();
		}
	}

	/** read the content of col contents */
	@Test
	public final void testGetColContentsWithBigIndex() {
		List<Object> contents = this.sheetReader.getColContents(1000);
		Assert.assertEquals(Collections.nCopies(this.sheetReader.getRowCount(), null), contents);
	}
	
	/** read the content of col contents */
	@Test(expected = IllegalArgumentException.class)
	public final void testGetColContentsWithNegativeIndex() {
		this.sheetReader.getColContents(-1);
		Assert.fail();
	}
	
	/** read the content of row 2 */
	@Test
	public final void testGetRow2() {
		try {
			final List</*@Nullable*/Object> rowContents = this.sheetReader
					.getRowContents(1);
			for (int i = 0; i < AbstractSpreadsheetReaderTest.ROW_2_CONTENTS
					.size(); i++)
				Assert.assertEquals(
						AbstractSpreadsheetReaderTest.ROW_2_CONTENTS.get(i),
						rowContents.get(i));
			Assert.assertEquals(AbstractSpreadsheetReaderTest.ROW_2_CONTENTS,
					rowContents);
		} catch (final IllegalArgumentException e) {
			Assert.fail();
		}
	}

	/** read the content of col contents */
	@Test
	public final void testGetRowContentsWithBigIndex() {
		List<Object> contents = this.sheetReader.getRowContents(1000);
		Assert.assertEquals(Collections.emptyList(), contents);
	}
	
	/** read the content of col contents */
	@Test(expected = IllegalArgumentException.class)
	public final void testGetRowContentsWithNegativeIndex() {
		this.sheetReader.getRowContents(-1);
		Assert.fail();
	}

	/** read the name and dimensions of the sheet */
	@Test
	public final void testSheetNameAndDimensions() {
		Assert.assertEquals(AbstractSpreadsheetReaderTest.SHEET_NAME,
				this.sheetReader.getName());
		Assert.assertEquals(AbstractSpreadsheetReaderTest.COL_C_CONTENTS.size(),
				this.sheetReader.getRowCount());
		Assert.assertEquals(AbstractSpreadsheetReaderTest.ROW_2_CONTENTS.size(),
				this.sheetReader.getCellCount(0));
	}

	/**
	 * @return properties for the test
	 */
	protected abstract TestProperties getProperties();
}