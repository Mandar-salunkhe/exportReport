package com.deliveryreview.request;

import java.util.ArrayList;
import java.util.List;

public class PartnerLeftOnesConsList {

	private List <PartnerLeftOnesConsList> inActivePartnerEcoSystemRows;

	public List<PartnerLeftOnesConsList> getInActivePartnerEcoSystemRows() {
		return inActivePartnerEcoSystemRows;
	}

	public void setInActivePartnerEcoSystemRows(List<PartnerLeftOnesConsList> inActivePartnerEcoSystemRows) {
		this.inActivePartnerEcoSystemRows = inActivePartnerEcoSystemRows;
	}

	@Override
	public String toString() {
		return "PartnerLeftOnesConsList [inActivePartnerEcoSystemRows=" + inActivePartnerEcoSystemRows + "]";
	}
	
	
	
}
