/**
 * 
 */
package main;

/**
 * @author jakem
 *
 */
public class Competitor {

private String name;
private String url;
private String select;
private String validation_select;
private String validation_type;
private String connection_type;

public String getName() {
    return name;
}

public void setName(String name) {
    this.name = name;
}

public String getUrl() {
    return url;
}

public void setUrl(String url) {
	this.url = url;
}
public String getSelect() {
    return select;
}

public void setSelect(String select) {
    this.select = select;
} 
public String getValidation_select() {
    return validation_select;
}

public void setValidation_select(String validation_select) {
    this.validation_select = validation_select;
} 

public String getValidation_type() {
    return validation_type;
}

public void setValidation_type(String validation_type) {
    this.validation_type = validation_type;
}
public String getConnection_type() {
    return connection_type;
}

public void setConnection_type(String connection_type) {
    this.connection_type = connection_type;
}

}
