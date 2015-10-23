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
import java.util.List;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

public abstract class SpreadsheetReaderTest {
	@Rule
	public TestName name = new TestName();
	protected SpreadsheetDocumentFactory factory;
	protected SpreadsheetDocumentReader sdr;

	protected SpreadsheetReader sr;

	/** set the test up */
	@Before
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL resourceURL = this.getClass().getResource(
					String.format("/VilleMTP_MTP_MonumentsHist.%s", this
							.getProperties().getExtension()));
			final InputStream inputStream = resourceURL.openStream();
			this.sdr = this.factory.openForRead(inputStream);
			Assert.assertEquals(1, this.sdr.getSheetCount());
			this.sr = this.sdr.getSpreadsheet(0);
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@After
	public void tearDown() {
		try {
			this.sdr.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testCellA1Content() {
		try {
			Assert.assertEquals("OBJECTID,N,10,0", this.sr.getCellContent(0, 0));
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	@Test
	public final void testCellA1IsAText() {
		try {
			Assert.assertEquals("OBJECTID,N,10,0", this.sr.getText(0, 0));
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellA1IsNotABoolean() {
		this.sr.getBoolean(0, 0);
		Assert.fail();
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellA1IsNotADate() {
		this.sr.getDate(0, 0);
		Assert.fail();
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellA1IsNotADouble() {
		this.sr.getDouble(0, 0);
		Assert.fail();
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellA1IsNotAFormula() {
		Assert.assertEquals(null, this.sr.getFormula(0, 0));
		Assert.fail();
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellA1IsNotAnInteger() {
		this.sr.getInteger(0, 0);
		Assert.fail();
	}

	@Test
	public final void testCellA2AsADouble() {
		try {
			Assert.assertEquals(Double.valueOf(170), this.sr.getDouble(1, 0));
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	@Test
	public final void testCellA2AsAnInteger() {
		try {
			Assert.assertEquals((Integer) 170, this.sr.getInteger(1, 0));
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	@Test
	public final void testCellA2Content() {
		try {
			Assert.assertEquals(170, this.sr.getCellContent(1, 0));
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellA2IsNotAFormula() {
		Assert.assertEquals(null, this.sr.getFormula(1, 0));
		Assert.fail();
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellA2IsNotAText() {
		Assert.assertEquals("170", this.sr.getText(1, 0));
		Assert.fail();
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellDateA2IsNotABoolean() {
		this.sr.getBoolean(1, 0);
		Assert.fail();
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testCellDateA2IsNotADate() {
		this.sr.getDate(1, 0);
		Assert.fail();
	}

	@Test
	public final void testGetColC() {
		try {
			Assert.assertEquals(
					Arrays.asList(
							"ETIQUETTE,C,254",
							"Immeuble dit « Hôtel Lefèvre »",
							"Gare",
							"Hôtel de Solas",
							"Portail",
							"Tour des Pins",
							"Hôtel de Fourques",
							"Hôtel de Bénézet",
							"Hôtel de Guidais",
							"Hôtel Haguenot",
							"Maison",
							"Ancien Logis du Chapeau rouge",
							"Hôtel Bardy",
							"Immeuble",
							"Ancien bureau d’octroi du Pont Juvénal",
							"Hôtel Verchant",
							"Hôtel de Querelles",
							"Hôtel de Cambacérès-Murles",
							"Hôtel de Ricard",
							"Hôtel de Campan",
							"Hôtel Montcalm",
							"Immeuble",
							"Hôtel de Lunas",
							"Hôtel des Carcassonne ou Hôtel de Gayon",
							"Jardin des plantes",
							"Ancien couvent des Ursulines, anciennes prisons (ex-caserne Grossetti)",
							"Ancienne prison",
							"Eglise des Pénitents Blancs",
							"Pont-aqueduc dit « Arceaux sur la Lironde »",
							"Hôtel Lamouroux",
							"Ancien cinéma Pathé",
							"Halle Castellane",
							"Vestiges de l'ancien ensemble cultuel hébraïque",
							"Temple protestant",
							"Faculté de Médecine et Musée d’anatomie",
							"Mas de Bagnères",
							"Ancienne église de Montels",
							"Ancien prieuré de Saint Pierre de  Montaubérou",
							"Domaine du château de Bonnier de la Mosson",
							"Citadelle",
							"Hôtel Bachy-du-Cayla",
							"Hôtel de Fizes",
							"Eglise Saint François de la Pierre Rouge, de l'Enclos Saint François",
							"Ancienne église des Cordeliers",
							"Evêché",
							"Hôtel Hortolès ou Ginestous",
							"Aqueduc des Arceaux ou de Pitot",
							"Fontaine des Licornes",
							"Porte de la Blanquerie",
							"Hôtel de Baudon et Mauny",
							"Fontaine de la Préfecture",
							"Hôtel de Fesquet",
							"Hôtel de Manse",
							"Hôtel d’Aurès",
							"Hôpital Général et Clinique Saint Charles",
							"Faculté de Médecine et Musée d’anatomie",
							"Ancien Hôtel de Ganges",
							"Fontaine des Trois Grâces",
							"Ensemble monumental de la Promenade du Peyrou",
							"Cathédrale Saint Pierre",
							"Eglise Sainte Croix de Celleneuve",
							"Eglise Notre-Dame-des-Tables",
							"Eglise Saint-Denis",
							"Eglise Sainte-Eulalie",
							"Hôtel de Mirman",
							"Chapelle de l'ancien couvent des Récollets",
							"hôtel de Grave, hôtel de Villarmois, hôtel de Noailles",
							"Hôtel Rey",
							"Hôtel de Sarret dit « de la Coquille »",
							"Hôtel des Trésoriers de la Bourse",
							"Hôtel de Varennes",
							"Vestiges de l'ancien ensemble cultuel hébraïque",
							"Jardins de la reine, Ancien Rectorat et Intendance Jardin des Plantes",
							"Ancien Hôtel de Belleval",
							"Hôtel de Joubert",
							"Hôtel d’Uston",
							"Ancienne maison ou couvent de la Miséricorde",
							"Hôtel de Bernard Duffau et ancien grand séminaire",
							"Hôtel Périer",
							"Hôtel de Boussugues",
							"Hôtel Pas de Beaulieu",
							"Hôtel de Castries",
							"Immeuble",
							"Ancienne église de la Visitation",
							"Eglise de Castelnau-le-Lez",
							"Eglise paroissiale Sainte Thérèse de Lisieux",
							"Hôtel de Claris",
							"Hôtel des Trésoriers de France (ou Hôtel du Lunaret)",
							"Escalier monumental", "Hôtel", "Hôtel Hostalier",
							"Hôtel de Saint Côme",
							"Hôtel de Magny ou Cabanes de Puimisson",
							"Hôtel de Castan", "Palais de Justice",
							"Château de Flaugergues", "Château Levat",
							"Château de la Mogère",
							"Château et Parc de la Piscine",
							"Ancien observatoire dit «Tour de la Babotte »",
							"Hôtel de Montferrier", "Hôtel de Griffy",
							"Hôtel Estorc",
							"Hôtel de la Société Royale des Sciences",
							"Hôtel des Vignes", "Hôtel Pomier-Layrargues",
							"Hôtel d'Hostalier", "Palais des Rois d’Aragon",
							"Hôtel de Roquemaure", "Hôtel d’Avèze",
							"Hôtel de Beaulac", "Hôtel Deydé",
							"Hôpital Général", "Hôtel de Saint-Félix",
							"Hôtel Lecourt", "Château d’Ô",
							"Collège des Ecossais"), this.sr.getColContents(2));
		} catch (final IllegalArgumentException e) {
			Assert.fail();
		}
	}

	@Test
	public final void testGetRow2() {
		try {
			final List<Object> list = Arrays.<Object> asList(
					Integer.valueOf(170),
					"Monuments historiques classés ou inscrits",
					"Immeuble dit « Hôtel Lefèvre »",
					"inscrit Inv. MH 19 11 1985", "En totalité",
					Double.valueOf(333.810638315),
					Double.valueOf(93.4774612976));
			final List<Object> rowContents = this.sr.getRowContents(1);
			for (int i = 0; i < list.size(); i++)
				Assert.assertEquals(list.get(i), rowContents.get(i));
			Assert.assertEquals(list, rowContents);
		} catch (final IllegalArgumentException e) {
			Assert.fail();
		}
	}

	@Test
	public final void testSheetNameAndDimensions() {
		Assert.assertEquals("Feuille1", this.sr.getName());
		Assert.assertEquals(117, this.sr.getRowCount());
		Assert.assertEquals(7, this.sr.getCellCount(0));
	}

	protected abstract TestProperties getProperties();
}