import React, { Component } from 'react';
import Header from '../header/header';
import Footer from '../footer/footer';
import Lipstick from '../../assets/images/landing_page_icons/icons/cross2.png';
import { Link } from 'react-router-dom';
import '../../assets/css/productAnalyzer.css';

class ProductAnalyzer extends Component {
    constructor(props) {
        super(props);
        this.state = {
            input: '',
            error: false
        }
    }

    handleSubmit(event) {
        event.preventDefault();
        if (this.state.input !== "") {
            let uriEncodedInput = encodeURIComponent(this.state.input);
            this.props.history.push('/product_analyzer_result/' + uriEncodedInput)
        } else {
            this.setState({ error: true })
        }
    }

    handleInput(event) {
        event.preventDefault();
        this.setState({
            input: event.target.value
        });
    }
    render() {
        return (
            <section>
                <Header history={this.props.history} />
                <div className="check-product-image-container">
                    <div className="logo-background-color">
                        <img className="check-product-logo" src={Lipstick} />
                    </div>
                    <div className="check-product-description">
                        Grab a personal care item, find the ingredients list online, and copy/paste the comma separated list here.
                    </div>
                    <form className="check-product-form-field">
                        <div className="check-product-input-container">
                            <textarea autoFocus onChange={this.handleInput.bind(this)} className="check-product-input-field" type="text" placeholder="ex. Titanium Dioxide, Water, Ethylhexyl Palmitate, Glycerin, Octyldodecanol, Silica, Pentylene Glycol"></textarea>
                        </div>
                        {this.state.error ? <div className="ingredients-error">Please enter ingredients to query.</div> : null}
                        <div className="check-product-button-container">
                            <button className="btn purple" onClick={this.handleSubmit.bind(this)}>Analyze</button>
                        </div>
                    </form>
                </div>
                <Footer />
            </section>
        )
    }
}
export default ProductAnalyzer